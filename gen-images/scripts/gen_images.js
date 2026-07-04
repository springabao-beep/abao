#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const http = require('http');
const https = require('https');
const os = require('os');

// ============================================================================
// Error Handling
// ============================================================================

function fail(message, statusCode = 1) {
  console.error(JSON.stringify({ ok: false, error: message }, null, 0));
  process.exit(statusCode);
}

// ============================================================================
// CLI Argument Parsing
// ============================================================================

function parseArgs() {
  const args = process.argv.slice(2);
  const params = {};

  // 多值字段：出现多次时聚合成数组
  const arrayFields = new Set(['image']);
  // 带 file: 前缀的 raw 值
  const rawFields = new Set(['image']);

  for (let i = 0; i < args.length; i++) {
    if (args[i].startsWith('--')) {
      const key = args[i].slice(2);
      // Convert kebab-case to camelCase
      const camelKey = key.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
      const value = args[++i];

      if (arrayFields.has(camelKey)) {
        if (!params[camelKey]) {
          params[camelKey] = [];
        }
        params[camelKey].push(value);
      } else if (rawFields.has(camelKey)) {
        params[camelKey] = value;
      } else {
        params[camelKey] = value;
      }
    }
  }

  return params;
}

// ============================================================================
// Configuration Loading
// ============================================================================

function loadClaudeSettings() {
  const settingsPath = path.join(os.homedir(), '.claude', 'settings.json');

  if (!fs.existsSync(settingsPath)) {
    throw new Error(`配置文件不存在: ${settingsPath}`);
  }

  try {
    const content = fs.readFileSync(settingsPath, 'utf8');
    const data = JSON.parse(content);
    const env = data.env || {};
    const baseUrl = env.ANTHROPIC_BASE_URL;
    const token = env.ANTHROPIC_AUTH_TOKEN;

    if (!baseUrl) {
      throw new Error('settings.json 中缺少 env.ANTHROPIC_BASE_URL');
    }
    if (!token) {
      throw new Error('settings.json 中缺少 env.ANTHROPIC_AUTH_TOKEN');
    }

    return {
      baseUrl: baseUrl.replace(/\/$/, ''),
      token
    };
  } catch (error) {
    if (error.message.includes('settings.json')) {
      throw error;
    }
    throw new Error(`读取 Claude settings.json 失败: ${error.message}`);
  }
}

function loadRuntimeSettings() {
  try {
    const settings = loadClaudeSettings();
    return { caller: 'claude', ...settings };
  } catch (error) {
    throw new Error(`Claude 配置读取失败: ${error.message}`);
  }
}

function buildApiUrl(caller, baseUrl, mode) {
  const endpoint = mode === 'generate' ? '/images/generations' : '/images/edits';
  if (caller === 'claude') {
    return `${baseUrl}/v1${endpoint}`;
  }
  return `${baseUrl}${endpoint}`;
}

// ============================================================================
// HTTP Requests
// ============================================================================

function postMultipart(url, token, fields, files) {
  return new Promise((resolve, reject) => {
    const boundary = '----ClaudeFormBoundary' + Date.now().toString(16);

    const urlObj = new URL(url);
    const isHttps = urlObj.protocol === 'https:';
    const requestLib = isHttps ? https : http;

    // 构建 multipart body
    const bodyParts = [];

    // 添加普通字段
    for (const [key, value] of Object.entries(fields)) {
      if (value !== undefined && value !== null) {
        bodyParts.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="${key}"\r\n\r\n${value}\r\n`, 'utf8'));
      }
    }

    // 添加文件（值可以是单文件或文件数组）
    for (const [key, fileValue] of Object.entries(files)) {
      const items = Array.isArray(fileValue) ? fileValue : [fileValue];
      for (const fileInfo of items) {
        const header = `--${boundary}\r\nContent-Disposition: form-data; name="${key}"; filename="${fileInfo.filename}"\r\nContent-Type: ${fileInfo.mime}\r\n\r\n`;
        bodyParts.push(Buffer.from(header, 'utf8'));
        bodyParts.push(fileInfo.data);
        bodyParts.push(Buffer.from('\r\n', 'utf8'));
      }
    }

    bodyParts.push(Buffer.from(`--${boundary}--\r\n`, 'utf8'));

    const body = Buffer.concat(bodyParts);

    const options = {
      hostname: urlObj.hostname,
      port: urlObj.port || (isHttps ? 443 : 80),
      path: urlObj.pathname + urlObj.search,
      method: 'POST',
      headers: {
        'Content-Type': `multipart/form-data; boundary=${boundary}`,
        'Authorization': `Bearer ${token}`,
        'Content-Length': body.length
      }
    };

    const req = requestLib.request(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) {
          try {
            const errorData = JSON.parse(body);
            const message = errorData.error?.message || errorData.message || body;
            reject(new Error(`接口调用失败: ${message}`));
          } catch {
            reject(new Error(`HTTP ${res.statusCode}: ${body.slice(0, 200)}`));
          }
        } else {
          try {
            resolve(JSON.parse(body));
          } catch (error) {
            reject(new Error(`解析响应失败: ${error.message}`));
          }
        }
      });
    });

    req.on('error', (err) => {
      reject(new Error(`网络请求失败: ${err.message}`));
    });

    req.write(body);
    req.end();
  });
}

function postJson(url, token, payload) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(payload);
    const urlObj = new URL(url);

    const isHttps = urlObj.protocol === 'https:';
    const requestLib = isHttps ? https : http;

    const options = {
      hostname: urlObj.hostname,
      port: urlObj.port || (isHttps ? 443 : 80),
      path: urlObj.pathname + urlObj.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`,
        'Content-Length': Buffer.byteLength(data)
      }
    };

    const req = requestLib.request(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) {
          try {
            const errorData = JSON.parse(body);
            const message = errorData.error?.message || errorData.message || body;
            reject(new Error(`接口调用失败: ${message}`));
          } catch {
            reject(new Error(`HTTP ${res.statusCode}`));
          }
        } else {
          try {
            resolve(JSON.parse(body));
          } catch (error) {
            reject(new Error(`解析响应失败: ${error.message}`));
          }
        }
      });
    });

    req.on('error', (err) => {
      reject(new Error(`网络请求失败: ${err.message}`));
    });

    req.write(data);
    req.end();
  });
}

// ============================================================================
// Image Processing
// ============================================================================

const MIME_TYPES = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.webp': 'image/webp',
  '.gif': 'image/gif'
};

function guessExtension(mime) {
  for (const [ext, type] of Object.entries(MIME_TYPES)) {
    if (type === mime) return ext.slice(1);
  }
  return 'png';
}

function chooseExtension(outputFormat, mime) {
  if (outputFormat) {
    const normalized = outputFormat.toLowerCase();
    if (normalized === 'jpeg') return 'jpg';
    return normalized;
  }
  if (mime) {
    return guessExtension(mime);
  }
  return 'png';
}

function fileToDataUrl(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`图片文件不存在: ${filePath}`);
  }

  const ext = path.extname(filePath).toLowerCase();
  const mime = MIME_TYPES[ext] || 'application/octet-stream';
  const data = fs.readFileSync(filePath);
  const base64 = data.toString('base64');
  return `data:${mime};base64,${base64}`;
}

function dataUrlToBuffer(dataUrl) {
  const match = dataUrl.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) {
    throw new Error('返回的 data URL 格式无效');
  }
  const mime = match[1];
  const buffer = Buffer.from(match[2], 'base64');
  return { mime, buffer };
}

function ensureOutputDir(customPath) {
  if (customPath) {
    const dir = path.dirname(customPath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    return dir;
  }
  // 使用用户运行脚本的当前工作目录
  const outputDir = path.join(process.cwd(), 'gen-images');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  return outputDir;
}

function formatLocalTimestamp() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hour = String(now.getHours()).padStart(2, '0');
  const minute = String(now.getMinutes()).padStart(2, '0');
  const second = String(now.getSeconds()).padStart(2, '0');
  return `${year}-${month}-${day}--${hour}-${minute}-${second}`;
}

function saveImages(imageEntries, outputFormat, baseName, outputPath) {
  const outputDir = ensureOutputDir(outputPath);
  const timestamp = baseName ? `_${Date.now()}` : formatLocalTimestamp();
  const paths = [];

  // 解析自定义输出路径
  let customBaseName = null;
  let customExt = null;
  if (outputPath) {
    const parsed = path.parse(outputPath);
    customBaseName = parsed.name;
    customExt = parsed.ext.slice(1); // 移除点号
  }

  for (let index = 0; index < imageEntries.length; index++) {
    const item = imageEntries[index];
    let ext = customExt || chooseExtension(outputFormat);
    let buffer;

    if (item.b64_json) {
      buffer = Buffer.from(item.b64_json, 'base64');
    } else if (item.url && item.url.startsWith('data:')) {
      const { mime, buffer: buf } = dataUrlToBuffer(item.url);
      ext = customExt || chooseExtension(outputFormat, mime);
      buffer = buf;
    } else {
      throw new Error('接口返回中未找到可保存的图片数据');
    }

    let filename;
    if (customBaseName) {
      // 使用自定义文件名
      if (imageEntries.length > 1) {
        filename = `${customBaseName}-${String(index + 1).padStart(2, '0')}.${ext}`;
      } else {
        filename = `${customBaseName}.${ext}`;
      }
    } else {
      // 使用默认文件名
      filename = baseName
        ? `${baseName}_edit${timestamp}.${ext}`
        : `${timestamp}-${String(index + 1).padStart(2, '0')}.${ext}`;
    }

    const filePath = path.join(outputDir, filename);
    fs.writeFileSync(filePath, buffer);
    paths.push(filePath);
  }

  return paths;
}

// ============================================================================
// Payload Building
// ============================================================================

function buildGenerationPayload(args) {
  if (!args.prompt) {
    throw new Error('缺少 prompt');
  }

  const payload = {
    model: args.model || 'gpt-image-2',
    prompt: args.prompt,
    response_format: 'b64_json',
    n: parseInt(args.n) || 1
  };

  const optionalFields = [
    'size', 'quality', 'background', 'outputFormat',
    'outputCompression', 'partialImages', 'moderation'
  ];

  for (const field of optionalFields) {
    if (args[field] !== undefined) {
      const payloadKey = field.replace(/([A-Z])/g, '_$1').toLowerCase();
      payload[payloadKey] = args[field];
    }
  }

  return payload;
}

function buildEditPayload(args) {
  if (!args.prompt) {
    throw new Error('缺少 prompt');
  }
  if (!args.image) {
    throw new Error('缺少要编辑的图片来源');
  }

  // 支持单图和多图
  const imageValues = Array.isArray(args.image) ? args.image : [args.image];
  const images = imageValues.map(v => {
    if (!v.startsWith('http://') && !v.startsWith('https://') && !v.startsWith('data:')) {
      return { image_url: fileToDataUrl(v) };
    }
    return { image_url: v };
  });

  const payload = {
    model: args.model || 'gpt-image-2',
    prompt: args.prompt,
    images,
    response_format: 'b64_json',
    stream: false,
    n: parseInt(args.n) || 1
  };

  if (args.mask) {
    let maskValue = args.mask;
    if (!maskValue.startsWith('http://') &&
        !maskValue.startsWith('https://') &&
        !maskValue.startsWith('data:')) {
      maskValue = fileToDataUrl(maskValue);
    }
    payload.mask = { image_url: maskValue };
  }

  const optionalFields = [
    'size', 'quality', 'background', 'outputFormat',
    'outputCompression', 'partialImages', 'moderation', 'inputFidelity'
  ];

  for (const field of optionalFields) {
    if (args[field] !== undefined) {
      const payloadKey = field.replace(/([A-Z])/g, '_$1').toLowerCase();
      payload[payloadKey] = args[field];
    }
  }

  return payload;
}

// ============================================================================
// Main
// ============================================================================

async function main() {
  try {
    const args = parseArgs();

    if (!args.mode) {
      throw new Error('缺少 --mode 参数');
    }

    const { caller, baseUrl, token } = loadRuntimeSettings();

    const url = buildApiUrl(caller, baseUrl, args.mode);
    let response;

    if (args.mode === 'generate') {
      const payload = buildGenerationPayload(args);
      response = await postJson(url, token, payload);
    } else if (args.mode === 'edit') {
      if (!args.prompt) {
        throw new Error('缺少 prompt');
      }
      if (!args.image) {
        throw new Error('缺少要编辑的图片来源');
      }

      // 统一为数组处理（单图兼容原 args.image 为字符串）
      const imageArgs = Array.isArray(args.image) ? args.image : [args.image];

      // 读取所有图片文件
      const imageFiles = [];
      for (const img of imageArgs) {
        const imagePath = img.startsWith('file:') ? img.slice(5) : img;
        if (!fs.existsSync(imagePath)) {
          throw new Error(`图片文件不存在: ${imagePath}`);
        }
        const ext = path.extname(imagePath).toLowerCase();
        const mime = MIME_TYPES[ext] || 'image/png';
        const imageData = fs.readFileSync(imagePath);
        imageFiles.push({
          filename: path.basename(imagePath),
          mime,
          data: imageData
        });
      }

      // 构建字段
      const fields = {
        model: args.model || 'gpt-image-2',
        prompt: args.prompt,
        response_format: 'b64_json',
        n: String(parseInt(args.n) || 1)
      };

      // 添加可选字段
      if (args.size) fields.size = args.size;
      if (args.quality) fields.quality = args.quality;
      if (args.background) fields.background = args.background;
      if (args.outputFormat) fields.output_format = args.outputFormat;
      if (args.outputCompression) fields.output_compression = args.outputCompression;
      if (args.partialImages) fields.partial_images = args.partialImages;
      if (args.moderation) fields.moderation = args.moderation;
      if (args.inputFidelity) fields.input_fidelity = args.inputFidelity;

      // 构建文件（多图时 image 为数组）
      const files = {
        image: imageFiles.length === 1 ? imageFiles[0] : imageFiles
      };

      // 添加 mask
      if (args.mask) {
        const maskPath = args.mask.startsWith('file:') ? args.mask.slice(5) : args.mask;
        if (!fs.existsSync(maskPath)) {
          throw new Error(`mask 文件不存在: ${maskPath}`);
        }
        const maskExt = path.extname(maskPath).toLowerCase();
        const maskMime = MIME_TYPES[maskExt] || 'image/png';
        const maskData = fs.readFileSync(maskPath);
        files.mask = {
          filename: path.basename(maskPath),
          mime: maskMime,
          data: maskData
        };
      }

      response = await postMultipart(url, token, fields, files);
    } else {
      throw new Error(`无效的 mode: ${args.mode}，必须是 generate 或 edit`);
    }

    const data = response.data;
    if (!Array.isArray(data) || data.length === 0) {
      throw new Error('接口返回中缺少 data');
    }

    const usedParams = {
      model: args.model || 'gpt-image-2',
      size: args.size,
      quality: args.quality,
      background: args.background,
      output_format: args.outputFormat || 'png',
      n: parseInt(args.n) || 1
    };

    if (args.mode === 'edit' && args.inputFidelity !== undefined) {
      usedParams.input_fidelity = args.inputFidelity;
    }

    // 编辑模式下使用第一张原图片名作为基础名
    let baseName = null;
    if (args.mode === 'edit' && args.image) {
      const firstImage = Array.isArray(args.image) ? args.image[0] : args.image;
      const firstPath = firstImage.startsWith('file:') ? firstImage.slice(5) : firstImage;
      baseName = path.basename(firstPath, path.extname(firstPath));
    }

    const paths = saveImages(data, args.outputFormat, baseName, args.output);

    console.log(JSON.stringify({
      ok: true,
      paths,
      used_params: usedParams
    }, null, 0));
  } catch (error) {
    fail(error.message);
  }
}

if (require.main === module) {
  main();
}

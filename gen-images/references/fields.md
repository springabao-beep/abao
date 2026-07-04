# gen-images 字段与交互规则

## 任务类型

- 文生图：调用 `POST /v1/images/generations`
- 改图：调用 `POST /v1/images/edits`

如果用户表达包含明确的图片来源（本地路径、URL、data URL）且语义是“修改图片 / 编辑图片 / 改图”，优先识别为改图。

## 必填字段

### 文生图
- `prompt`

如果缺少 `prompt`，先向用户追问，不要执行脚本。

### 改图
- `prompt`
- 图片来源

图片来源支持：
- 本地图片路径
- 图片 URL
- data URL
- **多张图片**：传多个 `--image` 参数（向 API 发送多个同名字段）

如果缺少图片来源，向用户明确提示：
1. 提供本地图片路径（可提供多张）
2. 提供图片 URL / data URL（可提供多张）

## 可选字段

- `model`：默认 `gpt-image-2`
- `response_format`：默认 `b64_json`
- `size`
- `quality`
- `background`
- `output_format`
- `output_compression`
- `partial_images`
- `n`：默认 `1`
- `moderation`
- `input_fidelity`（改图可用）
- `output`：完整输出路径（如 `D:\aaa\bbb.png`）

**注意**：`stream` 参数已移除，API 不支持流式响应。

## 自然语言映射

### size
- `1024x1024`、`1:1` -> `size=1024x1024`
- `1024x1536`、`3:4` -> `size=1024x1536`
- `1536x1024`、`4:3` -> `size=1536x1024`
- `2048x2048` -> `size=2048x2048`
- `3840x2160`、`16:9` -> `size=3840x2160`
- `2160x3840`、`9:16` -> `size=2160x3840`
- `4k横向`、`16:9` -> `size=3840x2160`
- `4k竖向`、`9:16` -> `size=2160x3840`
- `auto` -> `size=auto`
- 如果用户明确给出上述尺寸或比例，优先映射到对应 `size`

### quality
- `高清`、`高质量`、`高品质` -> `quality=high`
- `中等质量` -> `quality=medium`
- `低质量` -> `quality=low`

### background
- `透明背景` -> `background=transparent`
- `白色背景` -> `background=white`
- `黑色背景` -> `background=black`

### output_format
- `png` -> `output_format=png`
- `jpg` -> `output_format=jpg`
- `jpeg` -> `output_format=jpeg`
- `webp` -> `output_format=webp`

如果用户自然语言中明确要求保存格式，按要求保存；否则默认保存为 `png`。

### n
- `生成3张`、`来3张`、`输出3张` -> `n=3`
- 未指定时默认 `n=1`

### output
- `保存在 D:\aaa\bbb.png` -> `output=D:\aaa\bbb.png`
- `保存为 my-image.png` -> `output=my-image.png`
- `输出到 /tmp/result.png` -> `output=/tmp/result.png`
- `生成图片到文件 D:\test.png` -> `output=D:\test.png`
- 未指定时默认保存到 `./gen-images/` 目录

### 多图编辑（image 多次传参）
- 改图模式支持传入多张图片：`--image ./a.png --image ./b.jpg`
- 自然语言"用 A 和 B 做 X"、"把 A 融入 B"、"把 A 中的 X 换到 B 中的 Y"
- 识别到多图意图时，向用户确认图片来源数量

## 预期等待时长

所有请求统一等待 2-10 分钟。

在调用脚本前，**先向用户提示预期等待时长**：

```
正在生成图片，预计等待 2-10 分钟，请耐心等待...
```

## timeout 规则

所有图片生成/编辑请求的 Bash `timeout` 统一为 `600000`（10 分钟），不再按尺寸区分。

## 字段优先级

1. 用户明确写出的字段值
2. 用户自然语言中的明确要求
3. 脚本默认值

## 交互规则

### 文生图缺 prompt
使用简短追问：

`请补充图片提示词，例如你想生成什么画面。`

### 改图缺图片来源
使用简短追问：

`请提供要编辑的图片来源：1）本地路径 2）图片 URL / data URL`

### 改图缺 prompt
使用简短追问：

`请补充修改要求，例如你想把图片改成什么效果。`

## 成功输出格式

成功后向用户报告：
- `图片已生成, 图片路径: <路径>`
- `实际使用的关键参数: model=..., size=..., quality=..., output_format=..., n=...`

## 失败输出格式

失败后向用户报告：
- `生成失败: <简短错误原因>`

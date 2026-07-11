---
name: daily-log-report
description: >-
  配置日常工作日志系统与飞书日报推送。
  当用户要求搭建"每日工作日志"、"日报推送"、"记一下工作内容"、"周报自动生成"、"工作琐事记录"等系统时使用。
  包括：创建工作日志目录结构、设置日志模板、创建定时自动化将日志汇总为日报并通过飞书 Bot 推送。
agent_created: true
---

# Daily Log & Report System

## 概述

在项目的根目录下创建 `工作日志/` 文件夹，日常对话中对 WorkBuddy 说"记一下..."即可追加内容到当天的 `YYYY-MM-DD.md` 日志文件中。同时设置一个定时自动化任务，在指定时间读取当天的日志文件，提取三个板块（今日完成、遇到的问题、明日计划），整理成日报并通过飞书 Bot 推送给用户。

## 什么时候使用

当用户说以下这些时使用本技能：

- "设置每日工作日志系统"
- "帮我弄个日报推送"
- "记录工作琐事并生成周报/日报"
- "搞一个记一下的流程"
- "通过飞书推送日报"

## 实施步骤

### 1. 创建日志目录和模板

在工作目录下创建 `工作日志/` 目录，放入 README.md 说明文件和当天的日志模板。

日志模板格式：

```markdown
# 工作日志 YYYY-MM-DD (周X)

## ✅ 今日完成
- （填写完成事项）

## ⚠️ 遇到的问题
- （填写遇到的问题）

## 📋 明日计划
- （填写明天计划）
```

### 2. 检查飞书连接状态

```bash
cd <project-root> && lark-cli auth status --json --verify
```

如果已连接，获取用户的 `openId`（``ou_xxx``）用于推送。

如果未连接，使用 Split-Flow 流程引导用户授权。

### 3. 测试飞书消息发送通道

先用 Bot 身份发送一条测试消息，确认推送通道畅通：

```bash
cd <project-root> && lark-cli im +messages-send --as bot --user-id ou_xxx --markdown "测试消息"
```

如果 Bot 身份发送成功，后续自动化也用 `--as bot` 发送。

如果提示 `im:message.send_as_user` 权限缺失，优先尝试 Bot 身份；Bot 无法发送时再补充授权。

### 4. 创建定时自动化

使用 `automation_update` 工具创建自动化任务，关键配置：

- **name**: 自动化名称（可包含 emoji）
- **scheduleType**: "recurring"
- **rrule**: 每天定时，如 `FREQ=DAILY;BYHOUR=21;BYMINUTE=0`
- **cwds**: 项目根目录
- **status**: "ACTIVE"
- **prompt**: 包含以下步骤的指令：
  1. 读取 `工作日志/YYYY-MM-DD.md`（当天日期）
  2. 提取"✅ 今日完成"、"⚠️ 遇到的问题"、"📋 明日计划"三个板块
  3. 格式化为 Markdown 日报
  4. 通过 `lark-cli im +messages-send --as bot --user-id ou_xxx --markdown` 发送

### 5. 告知用户使用方法

告诉用户日常只需要对我说"记一下..."即可记录工作内容，系统会在每日指定时间自动推送日报。

## 注意事项

- **发送身份**: 优先用 `--as bot` 而不是 `--as user`，因为 bot 不需要 `send_as_user` 权限
- **日志目录**: 必须放在项目根目录下的 `工作日志/`
- **自动化 prompt**: 需要使用 `cd <project-path>` 确保路径正确，消息内容用 ANSI-C 引号 `$'...'` 包裹
- **飞书消息**: 使用 `--markdown` 格式发送格式化内容，不要用 `--text`

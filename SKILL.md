---
name: baidu-batch
description: 百度网盘批量转存+分享工具，基于 BaiduPanFilesTransfers 项目实现，支持批量转存分享链接并自动生成分享链接；当用户需要批量处理百度网盘资源转存分享、整理教育资源、打包分发学习资料时使用
metadata:
  openclaw:
    requires:
      env: ["BAIDU_COOKIE"]
    primaryEnv: "BAIDU_COOKIE"
---

# 百度网盘批量转存+分享自动化工具

## 任务目标
- 本 Skill 用于：批量转存多个网盘链接到指定文件夹，并自动创建分享链接
- 能力包含：批量转存、批量分享、链接检测、Excel 输出、自动去重
- 触发条件：用户提供网盘链接（文本或Excel格式）并要求批量处理

## 依赖等级
- 等级：L3
- 说明：需要 Python + requests 库，需配置百度网盘 Cookie

## 前置准备

### 环境初始化
```bash
pip install requests pandas openpyxl
```

### 配置网盘凭据
复制 `.env.example` 为 `.env`，填入以下信息：
- `BAIDU_COOKIE`：百度网盘 Cookie（从浏览器开发者工具获取）

**获取 Cookie 方法**：
1. 浏览器登录 [百度网盘主页](https://pan.baidu.com/)
2. 按 F12 打开开发者工具 → Network 标签
3. 刷新页面，点击以 `main` 开头的请求
4. 复制 Request Headers 中的 `Cookie` 字段（以 `BAIDUID` 开头）

> ⚠️ 必须获取 main 页面的 Cookie，其他页面的 Cookie 不完整会导致转存失败

## 操作步骤

### 步骤 1：输入数据准备

**方式一：文本格式**
用户提供网盘链接列表，每行一个链接：
```
百度教程 https://pan.baidu.com/s/1xxxxxx 6img
https://pan.baidu.com/s/1yyyyyy?pwd=abcd
```

**方式二：Excel 格式**
用户提供 Excel 文件，包含以下列：
- `链接`：百度网盘分享链接
- `名称`：资源名称（可选，留空则自动获取）

### 步骤 2：执行批量处理

调用主脚本处理：
```bash
python scripts/batch_share.py \
  --input <输入文件/文本> \
  --output outputs/tables/result.xlsx
```

**参数说明**：
- `--input`：输入文件路径或直接文本（每行一个链接）
- `--output`：输出 Excel 文件路径

### 步骤 3：输出结果

生成的 Excel 文件包含以下列：
- `资源名称`：文件或文件夹名称
- `网盘链接`：创建的新分享链接（自动附带提取码）
- `状态`：处理状态（成功/失败/跳过）
- `备注`：错误信息或跳过原因（如有）

## 功能特性

### 1. 转存去重
- **链接去重**：已处理过的链接自动跳过
- **文件去重**：目标目录「来自：分享」中已存在的文件自动跳过

### 2. 固定转存目录
- 转存目标目录固定为：**来自：分享**
- 如果目录不存在，自动创建

### 3. 自动创建分享
- **永久有效**：分享链接永久有效
- **随机提取码**：自动生成 4 位随机提取码
- **自动填充**：分享链接自动附带提取码，格式：`https://pan.baidu.com/s/xxx?pwd=xxxx`

## 使用示例

### 示例一：批量处理百度链接

**场景描述**：用户有 10 个百度网盘学习资料链接，需要批量转存并分享

**执行方式**：
```bash
python scripts/batch_share.py \
  --input "Python教程 https://pan.baidu.com/s/xxx
https://pan.baidu.com/s/yyy 提取码：abcd
https://pan.baidu.com/s/zzz?pwd=efgh" \
  --output outputs/tables/学习资源.xlsx
```

**预期输出**：
- 生成 `outputs/tables/学习资源.xlsx`
- 自动转存到「来自：分享」目录
- 每个资源生成带随机提取码的分享链接

### 示例二：从 Excel 读取并处理

**场景描述**：用户提供 `resources.xlsx`，包含 50 个网盘链接

**执行方式**：
```bash
python scripts/batch_share.py \
  --input resources.xlsx \
  --output outputs/tables/result.xlsx
```

## 注意事项

1. **转存限制**：单账号每日最多创建 300 个分享链接
2. **请求频率**：建议转存间隔 1-3 秒，总转存速度不要超过每分钟 60 条链接
3. **IP 限制**：连续转存 1000+ 链接可能触发 IP 封锁，必要时可重启网络更换 IP
4. **Cookie 失效**：如遇"链接访问次数过多"错误，需重新获取 Cookie
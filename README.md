# baidu-batch

百度网盘批量转存+分享工具，OpenClaw Skill

## 功能

- 批量转存百度网盘分享链接到指定文件夹
- 自动生成分享链接（永久有效 + 随机提取码）
- 链接检测与去重
- Excel 格式输出

## 依赖

- Python 3
- requests
- pandas
- openpyxl

## 配置

复制 `.env.example` 为 `.env`，填入 `BAIDU_COOKIE`（百度网盘 Cookie，从浏览器开发者工具获取）

详细说明见 SKILL.md

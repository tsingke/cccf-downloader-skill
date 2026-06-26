<div align="center">
  <img src="logo.svg" width="140" alt="CCCF Downloader Logo">
  <h1>CCCF Downloader · Skill</h1>
  <p><strong>《中国计算机学会通讯》期刊下载 & 跟踪 — Claude Code Skill</strong></p>

  <p>
    <img src="https://img.shields.io/badge/Python-3.6%2B-2563EB?style=flat&logo=python&logoColor=white" alt="Python">
    <img src="https://img.shields.io/badge/License-MIT-22C55E?style=flat" alt="License">
    <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-1E40AF?style=flat" alt="Platform">
    <img src="https://img.shields.io/github/stars/tsingke/cccf-downloader-skill?style=flat&logo=github" alt="Stars">
    <img src="https://img.shields.io/github/last-commit/tsingke/cccf-downloader-skill?style=flat&color=F59E0B" alt="Last Commit">
  </p>
  <p>
    <a href="#-安装">安装</a> •
    <a href="#-用法">用法</a> •
    <a href="#-跟踪表">跟踪表</a> •
    <a href="#-自动下载">自动下载</a>
  </p>
</div>

<br>

## 📦 安装

### 方式一：Claude Code 自动安装

```bash
# 在 Claude Code 中执行
/claude-add-skill cccf-download https://github.com/tsingke/cccf-downloader-skill
```

### 方式二：手动安装

```bash
git clone https://github.com/tsingke/cccf-downloader-skill.git ~/.claude/skills/cccf-download
pip3 install openpyxl  # 可选，用于 Excel 导出
python3 ~/.claude/skills/cccf-download/cccf_download.py --login
```

## 🚀 用法

### Claude Code 内触发

| 命令 | 说明 |
| :--- | :--- |
| `/cccf 2025 1-12` | 下载 2025 年全年 |
| `/cccf 2026 1` | 下载 2026 年第 1 期 |
| `/cccf update` | 智能检测并补全缺期 |
| `/cccf list` | 查看下载跟踪表 |
| `/cccf login` | 更新 CCF 登录 Cookie |

### 直接运行

```bash
python3 cccf_download.py --year 2025 --issues 1-12
python3 cccf_download.py --update
python3 cccf_download.py --list
python3 cccf_download.py --login
```

### 参数

| 参数 | 说明 |
| :--- | :--- |
| `--year YEAR` | 目标年份 |
| `--issues RANGE` | 期号范围如 `1-12` 或 `1,3,5` |
| `--list` / `-l` | 查看跟踪表 |
| `--update` / `-u` | 检测缺期并补全 |
| `--login` | 更新 Cookie |

## 📊 跟踪表

```
📊 CCCF 下载跟踪表

| 年份 | 类型   | 01 | 02 | 03 | 04 | 05 | 06 | 07 | 08 | 09 | 10 | 11 | 12 | 状态       |
| :--- | :---   | :- | :- | :- | :- | :- | :- | :- | :- | :- | :- | :- | :- | :---       |
| 2026 | 《计算》| ● | ● | ● | ● | ● | ● | — | — | — | — | — | — | 完整 6/6   |
| 2026 | 《快讯》| ○ | ○ | ○ | ○ | ○ | ○ | — | — | — | — | — | — | 无快讯     |
| 2025 | 《计算》| ● | ● | ● | ● | ● | ● | ● | ● | ● | ● | ● | ● | 完整 12/12 |
| 2025 | 《快讯》| ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ | 无快讯     |
```

同时生成 `CCCF_下载跟踪表.xlsx`（Excel 格式，带颜色编码）。

## ⏰ 定时自动下载

| 平台 | 配置 |
|------|------|
| macOS / Linux | `crontab -e` → `3 9 22 * * cd /path && python3 cccf_download.py --update` |
| Windows | 任务计划程序 → 每月 22 号 9:00 |

CCCF 通常每月 20 日上线，22 号执行留出缓冲。

## 📄 许可证

MIT

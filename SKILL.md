---
name: cccf-download
description: 下载《中国计算机学会通讯》(CCCF) 期刊 PDF，跟踪下载状态，定时自动补全
trigger: cccf
---

# CCCF Downloader — Claude Code Skill

自动从 CCF 数字图书馆下载《中国计算机学会通讯》(CCCF) 期刊整刊 PDF，
支持《计算》主刊 + 《快讯》简报版，生成下载跟踪表。

## 用法

### 在 Claude Code 中触发

```
/cccf 2025 1-6      下载 2025 年第 1-6 期
/cccf update        智能更新缺期
/cccf list          查看下载跟踪表
/cccf login         更新 CCF 登录 Cookie
```

### 直接运行

```bash
python3 cccf_download.py --year 2025 --issues 1-12
python3 cccf_download.py --update
python3 cccf_download.py --list
python3 cccf_download.py --login
```

## 工作流程

1. 通过 CCF 数字图书馆 Cookie 认证
2. 获取指定年份/期号的下载 ID
3. 下载《计算》主刊 PDF（如有《快讯》则一并下载）
4. 更新 Markdown 跟踪表 + Excel 导出
5. 支持 crontab / 任务计划程序定时自动执行

## 自动下载

| 平台 | 命令 |
|------|------|
| macOS / Linux | `3 9 22 * * cd /path && python3 cccf_download.py --update` |
| Windows | 任务计划程序 → 每月 22 号 9:00 → `python cccf_download.py --update` |

## 输出目录

`~/CCCF_Downloads/` — 所有 PDF + 跟踪表均保存在此。

## Cookie 有效期约 60 天

过期后运行 `--login` 重新获取。

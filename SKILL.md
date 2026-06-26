---
name: cccf-download
description: 下载《中国计算机学会通讯》(CCCF) 期刊 PDF，跟踪下载状态，定时自动补全
trigger: cccf
---

# CCCF Downloader

自动从 CCF 数字图书馆下载《中国计算机学会通讯》(CCCF) 期刊整刊 PDF，
支持《计算》主刊 + 《快讯》简报版，生成下载跟踪表。

## 用法

| 命令 | 说明 |
|------|------|
| `/cccf 2025 1-6` | 下载 2025 年第 1-6 期 |
| `/cccf update` | 智能更新缺期 |
| `/cccf list` | 查看下载跟踪表 |
| `/cccf login` | 更新 CCF 登录 Cookie |

## 工作流程

1. **登录**: `--login` 获取 Cookie → 有效期约 60 天
2. **下载**: 自动获取指定期号 PDF（《计算》+《快讯》）
3. **跟踪**: Markdown 表格 + Excel 双格式保存下载记录
4. **定时**: crontab / 任务计划 → 每月 22 号自动执行

## 输出

`~/CCCF_Downloads/` — PDF + 跟踪表均保存在此。

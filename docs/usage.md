# CCCF Downloader 详细使用文档

## 目录

- [CCF 登录](#ccf-登录)
- [Cookie 管理](#cookie-管理)
- [下载模式详解](#下载模式详解)
- [跟踪表说明](#跟踪表说明)
- [常见问题](#常见问题)

## CCF 登录

CCF 数字图书馆需要会员登录才能下载全文。

### 获取 Cookie

```bash
python3 cccf_download.py --login
```

1. 浏览器打开 [CCF 数字图书馆](https://dl.ccf.org.cn)
2. 使用 CCF 会员账号登录
3. 按 F12 打开开发者工具 → 切换到 Console 标签
4. 输入 `copy(document.cookie)` 并回车
5. 粘贴到终端提示处

### Cookie 有效期

| 项目 | 说明 |
|------|------|
| 有效期 | 约 60 天 |
| 过期表现 | `--update` 报错，或下载 0 字节文件 |
| 更新方式 | 重新运行 `--login` |

## Cookie 管理

Cookie 保存在 `cookies.json`（已在 `.gitignore` 中排除，不会提交到仓库）。

如需在多个设备间共享 Cookie，注意：
- Cookie 绑定 IP 和 Session 信息
- 复制到其他设备可能立即失效
- 建议各设备独立执行 `--login`

## 下载模式详解

### 模式 1：指定年份/期号

```bash
# 下载 2025 年 1-6 期
python3 cccf_download.py --year 2025 --issues 1-6

# 下载单期
python3 cccf_download.py --year 2026 --issues 6

# 下载多期（逗号分隔）
python3 cccf_download.py --year 2026 --issues 1,3,5
```

### 模式 2：智能更新

```bash
python3 cccf_download.py --update
```

自动扫描跟踪表，检测所有缺期并补全。适合：
- 每月定期执行
- 更换设备后补全
- 发现缺期后修复

### 模式 3：定时自动

通过 crontab 每月 22 号自动执行（详见主 README）。

## 跟踪表说明

### 数据来源

跟踪数据保存在 `~/.cccf_tracker.json`，每次下载后自动更新。

### Markdown 跟踪表

运行 `--list` 在终端中查看：

```
📊 CCCF 下载跟踪表

| 年份 | 类型   | 01 | 02 | 03 | 04 | 05 | 06 | 07 | 08 | 09 | 10 | 11 | 12 | 状态 |
| :--- | :---   | :- | :- | :- | :- | :- | :- | :- | :- | :- | :- | :- | :- | :-- |
| 2026 | 📖 计算 | ● | ● | ● | ● | ● | ● | — | — | — | — | — | — | 完整 |
|      | 📬 快讯 | ○ | ○ | ○ | ○ | ○ | ○ | — | — | — | — | — | — | 无 |
```

### Excel 跟踪表

每次下载后自动生成 `CCCF_下载跟踪表.xlsx`，位于输出目录下。

| 特性 | 说明 |
|------|------|
| 蓝色表头 | 冻结，滚动时可见 |
| 计算/快讯分列 | 每年两行，独立展示 |
| ✅ / ⬜ / —— | 分别对应已下载/缺期/未出版 |
| 汇总统计 | 下载总数 + 最近下载时间 |

### 图例

| 符号 | 含义 |
|:----:|------|
| ● | 已下载 |
| ○ | 缺期（待下载） |
| — | 未出版（未到月份） |

## 常见问题

### 下载的 PDF 是 0 字节？

Cookie 过期了，重新运行 `--login`。

### 如何更新到最新期？

```bash
python3 cccf_download.py --update
```

### 跟踪表的 Excel 文件在哪？

输出目录下的 `CCCF_下载跟踪表.xlsx`。

### crontab 没有执行？

检查日志：

```bash
cat ~/Workspace/Docs/CCF/.cccf_cron.log
```

常见原因：
- Cookie 过期（> 60 天）
- 网络不通
- crontab 未正确配置

### 如何修改输出目录？

编辑 `cccf_download.py` 顶部的 `OUTPUT_DIR` 变量。

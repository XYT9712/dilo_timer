# DILO 活动计时器

基于 Python 和 Streamlit 的 DILO 活动计时工具，支持快速记录不同动作编码的耗时、维护备注，并将记录导出为 Excel 工作簿。

## 功能

- 默认提供 `N1` 到 `N100` 编码，可添加或覆盖编码及动作说明，也可直接编辑当前编码的说明。
- 开始计时后切换编码，会自动保存上一段记录并从零开始计下一段；停止时保存当前段。
- 快速备注会随切换或停止时保存到对应记录，已有记录的备注也可单独编辑。
- 展示当前时长、累计时长和记录条数；查看记录表格。
- 导出含「计时记录」和「编码对照表」两个工作表的 `.xlsx` 文件。
- 重置计时器保留历史记录；清除全部数据需二次确认并输入 `123`。
- 适配窄屏和移动设备。

## 本地运行

需要 Python 3.10 或更高版本。

```bash
python -m venv .venv
```

Windows PowerShell：

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run dilo_timer.py
```

macOS / Linux：

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run dilo_timer.py
```

## 部署到 Streamlit Community Cloud

1. 将本仓库推送到 GitHub。
2. 登录 [Streamlit Community Cloud](https://share.streamlit.io/) 并授权访问此 GitHub 仓库。
3. 选择 **Create app**，指定仓库、分支和主文件 `dilo_timer.py`，然后部署。
4. 后续更新推送到所选分支后，Community Cloud 可重新部署应用。

GitHub 用于保存源代码；Streamlit Community Cloud 负责运行 Python 应用，GitHub Pages 本身不能运行 Streamlit。

## 数据存储说明

为避免在 Community Cloud 上把不同访问者的数据写到同一个共享文件，此版本将编码表、计时记录和当前计时保存在各自的 Streamlit 浏览器会话状态中。数据不会作为跨设备的云端数据库同步，也不应视为持久备份；结束或丢失会话后数据可能无法恢复。请定期使用「下载 Excel 工作簿」备份记录。若需要跨会话、跨设备的长期保存，应另行配置数据库或对象存储。

清除时要求输入的 `123` 是沿用原 HTML 的防误触步骤，不是账号认证或安全边界。

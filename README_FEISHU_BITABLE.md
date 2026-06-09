# 飞书多维表格上传说明

这个项目现在可以把清洗后的 KPI Excel 上传到飞书多维表格。建议上传用于展示的 sheet：

- `波次号去重统计`：每个波次一行，适合做明细展示。
- `拣货人汇总`：每个拣货人一行，适合做看板汇总。

## 1. 准备飞书多维表格

1. 在飞书开放平台给应用开通多维表格读写权限。
2. 在目标多维表格里添加该应用，并给编辑权限。
3. 复制多维表格的直接链接，建议是这种格式：

```text
https://xxx.feishu.cn/base/bascnxxxxxxxx?table=tblxxxxxxxx
```

如果链接是 `/wiki/`，需要进入实际多维表格页面后复制 `/base/` 链接，或手动配置 `FEISHU_APP_TOKEN` 和 `FEISHU_TABLE_ID`。

## 2. 配置本地环境

复制模板：

```powershell
Copy-Item feishu.env.example .env
```

然后编辑 `.env`：

```text
FEISHU_APP_ID=cli_xxx
FEISHU_APP_SECRET=xxx
FEISHU_BITABLE_URL=https://xxx.feishu.cn/base/bascnxxxxxxxx?table=tblxxxxxxxx
FEISHU_KPI_FILE=C:\Users\Administrator\-API-\output\kpi\delivery_kpi_20260608.xlsx
FEISHU_SHEET_NAME=波次号去重统计
FEISHU_CLEAR_TABLE=0
```

如果这个多维表格只负责展示当天数据，可以把 `FEISHU_CLEAR_TABLE=1`，每次上传前会先清空目标表，避免重复。

## 3. 本地测试

先只解析 Excel，不调用飞书接口：

```powershell
python feishu_bitable_upload.py "C:\Users\Administrator\-API-\output\kpi\delivery_kpi_20260608.xlsx" --sheet "波次号去重统计" --dry-run
```

确认列数、记录数正确后，再真实上传：

```powershell
python feishu_bitable_upload.py "C:\Users\Administrator\-API-\output\kpi\delivery_kpi_20260608.xlsx" --sheet "波次号去重统计"
```

如果要上传拣货人汇总：

```powershell
python feishu_bitable_upload.py "C:\Users\Administrator\-API-\output\kpi\delivery_kpi_20260608.xlsx" --sheet "拣货人汇总"
```

## 4. 调度中心里连起来

启动调度平台后，页面里会多一个任务：

```text
飞书上传清洗结果
```

它会读取 `.env` 里的 `FEISHU_KPI_FILE` 和 `FEISHU_SHEET_NAME`，把对应 sheet 上传到多维表格。确认稳定后，可以在 `scheduler_app/jobs.py` 里给这个任务加 `cron`，例如每天 10:00：

```python
"cron": "0 10 * * *",
```

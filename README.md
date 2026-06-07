# 领星 OMP/WMS 自动化

当前项目已支持：

- 创建 OMP 出库单 Excel 导出任务
- 轮询任务中心并下载导出结果
- 本地生成 `Track-Key`
- 清洗第一份出库 KPI Excel
- 端到端生成出库 KPI 工作簿

## 安装依赖

```powershell
python -m pip install -r requirements.txt
```

## 认证方式

脚本优先支持 Playwright 导出的 `storage_state.json`。把文件放到项目根目录：

```text
storage_state.json
```

也可以用环境变量：

```powershell
$env:OMP_TOKEN="你的 OMP token"
$env:OMP_COOKIE="可选 cookie"
```

## 一次性生成出库 KPI

默认使用最近一个已完成的工作日 KPI 日期：

```powershell
python run_delivery_kpi_once.py --storage-state storage_state.json --fetch-picker-names
```

指定 KPI 日期：

```powershell
python run_delivery_kpi_once.py --storage-state storage_state.json --kpi-date 2026-06-05 --fetch-picker-names
```

输出文件：

- 原始导出：`output/raw/delivery_kpi_raw_YYYYMMDD_<taskId>.xlsx`
- 清洗结果：`output/kpi/delivery_kpi_YYYYMMDD.xlsx`

## KPI 时间规则

- 周一到周四：前一天 `22:30:00` 到当天 `22:30:00`
- 周五：前一天 `22:30:00` 到当天 `23:59:59`
- 清洗时会从出库单中去除前一天 `22:30:00` 到当天 `10:30:00` 的拣货时间数据
- 保留 `出库单` sheet
- 生成 `拣货时间去除` 和 `波次号去重统计` sheet

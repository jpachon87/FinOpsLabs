---
layout: lab
title: "Demostración 5: Construir un tablero FinOps básico"
permalink: /lab5/lab5/
images_base: /labs/lab5/img
duration: "60 minutos"
objective:
  - Construir un tablero FinOps en Excel que integre KPIs de costo, tendencias, presupuesto, forecast, Allocation y gestión de anomalías usando datos generados en las prácticas anteriores.
prerequisites:
  - Equipo con Windows 11 y Visual Studio Code instalado.
  - Python 3.12 o superior disponible desde la terminal integrada de Visual Studio Code.
  - Paquetes pandas, numpy, matplotlib y openpyxl instalados.
  - Haber completado las prácticas 3 y 4 o disponer de sus archivos de salida.
  - Conocimiento básico de Budgeting, Forecasting, Allocation, Anomaly Management y KPIs FinOps.
introduction:
  - En esta demostración integrarás los artefactos generados previamente para construir un tablero FinOps funcional de NovaTech SRL. Consolidarás KPIs, Budget vs Actual, forecast de los próximos tres meses, anomalías y métricas de Allocation en un único archivo Excel.
slug: lab5
lab_number: 5
final_result: >
  Al finalizar tendrás un archivo novatech_finops_dashboard_v1.xlsx con cinco hojas: Executive Summary, Cost & Trend, Budget vs Actual, Forecast y Anomaly Log.
notes:
  - Todos los archivos se almacenan dentro de la carpeta finops-course del Escritorio de Windows.
  - Todos los importes monetarios se expresan en EUR.
  - Los thresholds RAG y los presupuestos son políticas internas de NovaTech para esta demostración.
  - El forecast usa una regresión lineal simplificada con fines didácticos y no sustituye un proceso de forecasting empresarial completo.
  - La banda mostrada alrededor del forecast representa ±1.96 desviaciones estándar de los residuos; no debe interpretarse como un intervalo de confianza estadístico completo.
references:
  - text: "FinOps KPIs & Benchmarking"
    url: https://www.finops.org/framework/capabilities/kpis-benchmarking/
  - text: "FinOps Budgeting"
    url: https://www.finops.org/framework/capabilities/budgeting/
  - text: "FinOps Forecasting"
    url: https://www.finops.org/framework/capabilities/forecasting/
  - text: "FinOps Anomaly Management"
    url: https://www.finops.org/framework/capabilities/anomaly-management/
prev: /lab4/lab4/
next: /lab6/lab6/
---

---

<!-- Aquí comienzan las instrucciones paso a paso de la práctica -->

## 📥 Tarea 1. Consolidar los datos FinOps — 12 min

En esta tarea prepararás los insumos del tablero y calcularás los KPIs principales utilizando los datos de Billing, tendencias, Allocation y presupuesto.

### Tarea 1.1. Preparar los archivos de entrada

- {% include step_label.html %} Abre Visual Studio Code y una terminal integrada de PowerShell.

  > **Nota:** Usa **Terminal → New Terminal** y confirma que el perfil activo sea PowerShell.
  {: .lab-note .info .compact}

  > **Salida esperada:** VS Code muestra una terminal PowerShell lista para ejecutar comandos.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la carpeta del laboratorio y crea los directorios necesarios.

  > **Importante:** La variable `$LabRoot` permite trabajar sin depender del nombre del usuario de Windows.
  {: .lab-note .important .compact}

  ```powershell
  $LabRoot = Join-Path ([Environment]::GetFolderPath('Desktop')) 'finops-course'
  New-Item -ItemType Directory -Force -Path `
    "$LabRoot\data\raw", `
    "$LabRoot\data\processed", `
    "$LabRoot\outputs\lab5", `
    "$LabRoot\scripts" | Out-Null
  Set-Location $LabRoot
  ```

  > **Salida esperada:** Existen `data\raw`, `data\processed`, `outputs\lab5` y `scripts`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la URL del archivo de presupuesto y descárgalo.

  ```powershell
  $BudgetUrl = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_team_budgets_eur.csv"

  Invoke-WebRequest `
    -Uri $BudgetUrl `
    -OutFile "$LabRoot\data\raw\novatech_team_budgets_eur.csv"
  ```

  > **Salida esperada:** Se crea `data\raw\novatech_team_budgets_eur.csv`.
  {: .lab-note .output .compact}

### Tarea 1.2. Cargar datos y calcular KPIs

- {% include step_label.html %} Crea `scripts\build_finops_dashboard.py` y agrega la carga de los archivos de entrada.

  > **Importante:** Esta práctica reutiliza los archivos producidos en las prácticas 3 y 4.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  from pathlib import Path
  import pandas as pd
  import numpy as np
  import matplotlib.pyplot as plt
  from openpyxl import Workbook, load_workbook
  from openpyxl.styles import Font, PatternFill, Alignment
  from openpyxl.chart import BarChart, LineChart, Reference

  ROOT = Path.home() / "Desktop" / "finops-course"
  RAW = ROOT / "data" / "raw"
  PROCESSED = ROOT / "data" / "processed"
  OUTPUTS = ROOT / "outputs" / "lab5"

  OUTPUTS.mkdir(parents=True, exist_ok=True)

  billing = pd.read_csv(
      RAW / "novatech_billing_4months_eur.csv",
      parse_dates=["ChargePeriodStart"]
  )

  trend_xlsx = PROCESSED / "novatech_trend_report.xlsx"
  allocation_xlsx = PROCESSED / "novatech_allocation_report.xlsx"
  budget_csv = RAW / "novatech_team_budgets_eur.csv"

  budgets = pd.read_csv(budget_csv)

  trend_summary = pd.read_excel(trend_xlsx, sheet_name="summary")
  anomalies = pd.read_excel(trend_xlsx, sheet_name="anomalies")
  growth = pd.read_excel(trend_xlsx, sheet_name="growth")
  allocation_metrics = pd.read_excel(allocation_xlsx, sheet_name="summary")

  print(f"Billing: {len(billing):,} registros")
  print(f"Budgets: {len(budgets)} equipos")
  print(f"Anomalías: {len(anomalies)}")
  ```

  > **Salida esperada:** Se cargan Billing, Budget, Trend y Allocation sin errores.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega el cálculo de los seis KPIs principales.

  > **Nota:** Todos los KPIs se derivan de datos generados en prácticas anteriores; no se introducen porcentajes arbitrarios.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo de** `print(f"Anomalías: {len(anomalies)}")`.
  {: .lab-note .important .compact}

  ```python
  billing["Month"] = billing["ChargePeriodStart"].dt.to_period("M").astype(str)

  monthly = (
      billing.groupby("Month", as_index=False)
      .agg(EffectiveCostEUR=("EffectiveCost", "sum"))
  )

  jan_cost = monthly.loc[
      monthly["Month"] == "2024-01",
      "EffectiveCostEUR"
  ].iloc[0]

  dec_cost = monthly.loc[
      monthly["Month"] == "2023-12",
      "EffectiveCostEUR"
  ].iloc[0]

  mom_pct = (jan_cost - dec_cost) / dec_cost * 100

  top_growth_service = growth.iloc[0]["ServiceName"]
  anomaly_count = len(anomalies)

  alloc_map = dict(
      zip(allocation_metrics["Metric"], allocation_metrics["Value"])
  )

  direct_pct = float(alloc_map["DirectAllocationPct"])
  residual_pct = float(alloc_map["ResidualUnallocatedPct"])

  kpis = pd.DataFrame({
      "KPI": [
          "EffectiveCost Enero",
          "Variación MoM",
          "Servicio con mayor crecimiento",
          "Anomalías detectadas",
          "Direct Allocation",
          "Residual Unallocated"
      ],
      "Value": [
          f"€{jan_cost:,.2f}",
          f"{mom_pct:+.1f}%",
          top_growth_service,
          str(anomaly_count),
          f"{direct_pct:.1f}%",
          f"{residual_pct:.1f}%"
      ]
  })

  print(kpis.to_string(index=False))
  ```

  > **Salida esperada:** Se muestran seis KPIs calculados con Billing, Trend y Allocation.
  {: .lab-note .output .compact}

{% assign results = site.data.task-results[page.slug].results %}
{% capture r1 %}{{ results[0] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r1 %}

---

## 💶 Tarea 2. Construir Budget vs Actual y Forecast — 15 min

En esta tarea calcularás la ejecución presupuestaria por equipo y generarás una proyección simplificada para febrero, marzo y abril de 2024.

### Tarea 2.1. Calcular Budget vs Actual

- {% include step_label.html %} Agrega el cálculo del gasto de enero por equipo usando el dataset de Allocation.

  > **Nota:** El presupuesto mensual utilizado pertenece a NovaTech y se distribuye por equipo.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo de** `print(kpis.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  allocation_raw = pd.read_csv(
      RAW / "novatech_allocation_4months_eur.csv",
      parse_dates=["ChargePeriodStart"]
  )

  jan_alloc = allocation_raw[
      allocation_raw["ChargePeriodStart"].dt.to_period("M").astype(str) == "2024-01"
  ]

  actual_by_team = (
      jan_alloc[jan_alloc["TagTeam"].notna()]
      .groupby("TagTeam", as_index=False)
      .agg(ActualEUR=("EffectiveCost", "sum"))
      .rename(columns={"TagTeam": "Team"})
  )

  budget_vs_actual = budgets.merge(
      actual_by_team,
      on="Team",
      how="left"
  ).fillna({"ActualEUR": 0})

  budget_vs_actual["ExecutionPct"] = (
      budget_vs_actual["ActualEUR"]
      / budget_vs_actual["MonthlyBudgetEUR"]
      * 100
  )

  budget_vs_actual["VarianceEUR"] = (
      budget_vs_actual["ActualEUR"]
      - budget_vs_actual["MonthlyBudgetEUR"]
  )
  ```

  > **Salida esperada:** Cada equipo muestra Budget, Actual, Execution % y Variance.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega la política RAG de NovaTech para Budget.

  > **Importante:** Los thresholds son política interna de NovaTech para esta demostración.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo de** `budget_vs_actual["VarianceEUR"] = (...)` y después del cálculo de `VarianceEUR`.
  {: .lab-note .important .compact}

  ```python
  def budget_status(pct):
      if pct > 100:
          return "RED"
      if pct > 80:
          return "AMBER"
      return "GREEN"

  budget_vs_actual["RAG"] = (
      budget_vs_actual["ExecutionPct"]
      .apply(budget_status)
  )

  print(
      budget_vs_actual[
          ["Team", "MonthlyBudgetEUR", "ActualEUR", "ExecutionPct", "RAG"]
      ].to_string(index=False)
  )
  ```

  > **Salida esperada:** Cada equipo queda marcado como GREEN, AMBER o RED según su ejecución presupuestaria.
  {: .lab-note .output .compact}

### Tarea 2.2. Generar forecast febrero–abril 2024

- {% include step_label.html %} Agrega la regresión lineal sobre el EffectiveCost diario.

  > **Nota:** El modelo lineal es didáctico y utiliza aproximadamente 123 días históricos.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo del bloque que imprime** `budget_vs_actual`.
  {: .lab-note .important .compact}

  ```python
  daily = (
      billing.groupby("ChargePeriodStart", as_index=False)
      .agg(EffectiveCostEUR=("EffectiveCost", "sum"))
      .sort_values("ChargePeriodStart")
  )

  daily["DayNum"] = (
      daily["ChargePeriodStart"] - daily["ChargePeriodStart"].min()
  ).dt.days

  x = daily["DayNum"].to_numpy()
  y = daily["EffectiveCostEUR"].to_numpy()

  slope, intercept = np.polyfit(x, y, 1)
  y_hist = slope * x + intercept

  residual_std = np.std(y - y_hist)

  future_dates = pd.date_range(
      start=daily["ChargePeriodStart"].max() + pd.Timedelta(days=1),
      end="2024-04-30",
      freq="D"
  )

  future_x = (
      future_dates - daily["ChargePeriodStart"].min()
  ).days.to_numpy()

  forecast_values = slope * future_x + intercept
  ```

  > **Salida esperada:** Se calculan valores diarios de forecast desde febrero hasta abril de 2024.
  {: .lab-note .output .compact}

- {% include step_label.html %} Construye el resumen mensual del forecast.

  > **Importante:** La banda ±1.96σ es indicativa y no representa un intervalo de confianza estadístico completo.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo de** `forecast_values = slope * future_x + intercept`.
  {: .lab-note .important .compact}

  ```python
  forecast_daily = pd.DataFrame({
      "Date": future_dates,
      "ForecastEUR": forecast_values,
      "UpperBandEUR": forecast_values + 1.96 * residual_std,
      "LowerBandEUR": forecast_values - 1.96 * residual_std
  })

  forecast_daily["Month"] = (
      forecast_daily["Date"].dt.to_period("M").astype(str)
  )

  forecast_monthly = (
      forecast_daily.groupby("Month", as_index=False)
      .agg(
          ForecastEUR=("ForecastEUR", "sum"),
          LowerBandEUR=("LowerBandEUR", "sum"),
          UpperBandEUR=("UpperBandEUR", "sum")
      )
  )

  print(forecast_monthly.to_string(index=False))
  ```

  > **Salida esperada:** Se muestran forecast mensual para `2024-02`, `2024-03` y `2024-04`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Genera la gráfica histórica y de forecast.

  > **Nota:** La imagen se utilizará como referencia visual en el dashboard.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo de** `print(forecast_monthly.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  plt.figure(figsize=(12, 6))

  plt.plot(
      daily["ChargePeriodStart"],
      daily["EffectiveCostEUR"],
      label="Actual diario"
  )

  plt.plot(
      future_dates,
      forecast_values,
      label="Forecast"
  )

  plt.fill_between(
      future_dates,
      forecast_daily["LowerBandEUR"],
      forecast_daily["UpperBandEUR"],
      alpha=0.15
  )

  plt.xlabel("Fecha")
  plt.ylabel("EffectiveCost diario (EUR)")
  plt.legend()
  plt.grid(alpha=0.3)
  plt.tight_layout()

  forecast_png = OUTPUTS / "forecast_feb_apr_2024.png"
  plt.savefig(forecast_png, dpi=150, bbox_inches="tight")
  plt.close()

  print(f"Gráfica creada: {forecast_png}")
  ```

  > **Salida esperada:** Se crea `outputs\lab5\forecast_feb_apr_2024.png`.
  {: .lab-note .output .compact}

{% capture r2 %}{{ results[1] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r2 %}

---

## 🚨 Tarea 3. Integrar Anomaly Management — 12 min

En esta tarea transformarás las anomalías detectadas en la Práctica 3 en un log operativo con prioridad, responsable y estado.

### Tarea 3.1. Preparar el Anomaly Log

- {% include step_label.html %} Agrega una política de prioridad para las anomalías detectadas.

  > **Importante:** Los niveles P1/P2/P3 y sus SLA son política interna de NovaTech.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo de** `print(f"Gráfica creada: {forecast_png}")`.
  {: .lab-note .important .compact}

  ```python
  def anomaly_priority(cost):
      if cost >= 5000:
          return "P1"
      if cost >= 1000:
          return "P2"
      return "P3"

  def anomaly_sla(priority):
      return {
          "P1": "24 h",
          "P2": "72 h",
          "P3": "5 días hábiles"
      }[priority]

  anomaly_log = anomalies.copy()

  anomaly_log["Priority"] = (
      anomaly_log["EffectiveCostEUR"]
      .apply(anomaly_priority)
  )

  anomaly_log["SLA"] = (
      anomaly_log["Priority"]
      .apply(anomaly_sla)
  )
  ```

  > **Salida esperada:** Cada anomalía queda clasificada como P1, P2 o P3.
  {: .lab-note .output .compact}

- {% include step_label.html %} Asigna responsables por servicio.

  > **Nota:** El owner se determina según el tipo de servicio y la estructura organizacional ficticia de NovaTech.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo de** `anomaly_log["SLA"] = (...)` y después de calcular la columna `SLA`.
  {: .lab-note .important .compact}

  ```python
  owner_map = {
      "Amazon EC2": "carlos.ruiz@novatech.io",
      "AWS Lambda": "carlos.ruiz@novatech.io",
      "Azure Virtual Machines": "ana.lopez@novatech.io",
      "Azure SQL Database": "diego.martinez@novatech.io",
      "BigQuery": "diego.martinez@novatech.io",
      "Compute Engine": "laura.garcia@novatech.io"
  }

  anomaly_log["Owner"] = (
      anomaly_log["ServiceName"]
      .map(owner_map)
      .fillna("finops@novatech.io")
  )

  anomaly_log["Status"] = "Pending Triage"
  anomaly_log["Action"] = ""
  anomaly_log["ResolvedAt"] = ""

  print(
      anomaly_log[
          ["Week", "ServiceName", "EffectiveCostEUR", "Priority", "Owner", "Status"]
      ].to_string(index=False)
  )
  ```

  > **Salida esperada:** Cada anomalía muestra prioridad, owner y estado inicial `Pending Triage`.
  {: .lab-note .output .compact}

### Tarea 3.2. Calcular métricas de gestión de anomalías

- {% include step_label.html %} Calcula el resumen operativo del Anomaly Log.

  > **Nota:** Estas métricas permiten medir carga de trabajo y severidad del proceso de triage.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo del bloque que imprime** `anomaly_log`.
  {: .lab-note .important .compact}

  ```python
  anomaly_metrics = pd.DataFrame({
      "Metric": [
          "TotalAnomalies",
          "P1",
          "P2",
          "P3",
          "PendingTriage"
      ],
      "Value": [
          len(anomaly_log),
          (anomaly_log["Priority"] == "P1").sum(),
          (anomaly_log["Priority"] == "P2").sum(),
          (anomaly_log["Priority"] == "P3").sum(),
          (anomaly_log["Status"] == "Pending Triage").sum()
      ]
  })

  print(anomaly_metrics.to_string(index=False))
  ```

  > **Salida esperada:** Se muestran cantidad total de anomalías y distribución por prioridad.
  {: .lab-note .output .compact}

- {% include step_label.html %} Guarda temporalmente el Anomaly Log en `outputs\lab5`.

  > **Nota:** El CSV facilita revisar los datos antes de integrarlos al Excel final.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo de** `print(anomaly_metrics.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  anomaly_csv = OUTPUTS / "anomaly_log.csv"
  anomaly_log.to_csv(anomaly_csv, index=False)

  print(f"Anomaly Log creado: {anomaly_csv}")
  ```

  > **Salida esperada:** Se crea `outputs\lab5\anomaly_log.csv`.
  {: .lab-note .output .compact}

{% capture r3 %}{{ results[2] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r3 %}

---

## 📊 Tarea 4. Construir el dashboard Excel — 16 min

En esta tarea generarás automáticamente un archivo Excel con cinco hojas y elementos visuales para presentar la información a distintas audiencias.

### Tarea 4.1. Crear Executive Summary y Cost & Trend

- {% include step_label.html %} Agrega la creación del workbook y la hoja `Executive Summary`.

  > **Importante:** El dashboard se genera automáticamente para mantener consistencia y reducir errores manuales.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo de** `print(f"Anomaly Log creado: {anomaly_csv}")`.
  {: .lab-note .important .compact}

  ```python
  dashboard_xlsx = OUTPUTS / "novatech_finops_dashboard_v1.xlsx"

  wb = Workbook()
  ws = wb.active
  ws.title = "Executive Summary"

  headers = ["KPI", "Value"]
  for col, value in enumerate(headers, start=1):
      cell = ws.cell(row=1, column=col, value=value)
      cell.font = Font(bold=True)

  for row_idx, row in enumerate(kpis.itertuples(index=False), start=2):
      ws.cell(row=row_idx, column=1, value=row.KPI)
      ws.cell(row=row_idx, column=2, value=row.Value)

  ws["D1"] = "NovaTech FinOps Dashboard"
  ws["D1"].font = Font(bold=True, size=16)
  ```

  > **Salida esperada:** Se crea la hoja `Executive Summary` con los seis KPIs.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega la hoja `Cost & Trend`.

  > **Nota:** Esta hoja presenta histórico mensual y servicios con mayor crecimiento.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo de** `ws["D1"].font = Font(bold=True, size=16)`.
  {: .lab-note .important .compact}

  ```python
  ws_cost = wb.create_sheet("Cost & Trend")

  for c, name in enumerate(["Month", "EffectiveCostEUR"], start=1):
      ws_cost.cell(row=1, column=c, value=name).font = Font(bold=True)

  for r, row in enumerate(monthly.itertuples(index=False), start=2):
      ws_cost.cell(row=r, column=1, value=row.Month)
      ws_cost.cell(row=r, column=2, value=float(row.EffectiveCostEUR))

  start_row = len(monthly) + 4

  ws_cost.cell(row=start_row, column=1, value="ServiceName").font = Font(bold=True)
  ws_cost.cell(row=start_row, column=2, value="GrowthPct").font = Font(bold=True)

  for r, row in enumerate(growth.head(5).itertuples(index=False), start=start_row + 1):
      ws_cost.cell(row=r, column=1, value=row.ServiceName)
      ws_cost.cell(row=r, column=2, value=float(row.GrowthPct))
  ```

  > **Salida esperada:** La hoja contiene gasto mensual y top 5 servicios por crecimiento.
  {: .lab-note .output .compact}

### Tarea 4.2. Crear Budget, Forecast y Anomaly Log

- {% include step_label.html %} Agrega la hoja `Budget vs Actual`.

  > **Nota:** El estado RAG proviene de la política interna de NovaTech definida anteriormente.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo del bloque que carga el top 5 de crecimiento en** `ws_cost`.
  {: .lab-note .important .compact}

  ```python
  ws_budget = wb.create_sheet("Budget vs Actual")

  budget_cols = [
      "Team", "MonthlyBudgetEUR", "ActualEUR",
      "ExecutionPct", "VarianceEUR", "RAG", "Owner"
  ]

  for c, name in enumerate(budget_cols, start=1):
      ws_budget.cell(row=1, column=c, value=name).font = Font(bold=True)

  for r, row in enumerate(budget_vs_actual.itertuples(index=False), start=2):
      ws_budget.cell(row=r, column=1, value=row.Team)
      ws_budget.cell(row=r, column=2, value=float(row.MonthlyBudgetEUR))
      ws_budget.cell(row=r, column=3, value=float(row.ActualEUR))
      ws_budget.cell(row=r, column=4, value=float(row.ExecutionPct))
      ws_budget.cell(row=r, column=5, value=float(row.VarianceEUR))
      ws_budget.cell(row=r, column=6, value=row.RAG)
      ws_budget.cell(row=r, column=7, value=row.Owner)
  ```

  > **Salida esperada:** Se crea la hoja con Budget, Actual, Variance y RAG por equipo.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega las hojas `Forecast` y `Anomaly Log`.

  > **Nota:** El dashboard conserva datos analíticos y operativos en hojas separadas.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **debajo del bloque que llena** `ws_budget`.
  {: .lab-note .important .compact}

  ```python
  ws_forecast = wb.create_sheet("Forecast")

  for c, name in enumerate(forecast_monthly.columns, start=1):
      ws_forecast.cell(row=1, column=c, value=name).font = Font(bold=True)

  for r, row in enumerate(forecast_monthly.itertuples(index=False), start=2):
      ws_forecast.cell(row=r, column=1, value=row.Month)
      ws_forecast.cell(row=r, column=2, value=float(row.ForecastEUR))
      ws_forecast.cell(row=r, column=3, value=float(row.LowerBandEUR))
      ws_forecast.cell(row=r, column=4, value=float(row.UpperBandEUR))

  ws_anom = wb.create_sheet("Anomaly Log")

  anomaly_cols = [
      "Week", "ServiceName", "ZScore", "EffectiveCostEUR",
      "Type", "Priority", "SLA", "Owner", "Status", "Action", "ResolvedAt"
  ]

  for c, name in enumerate(anomaly_cols, start=1):
      ws_anom.cell(row=1, column=c, value=name).font = Font(bold=True)

  for r, row in enumerate(anomaly_log.itertuples(index=False), start=2):
      values = [
          row.Week,
          row.ServiceName,
          row.ZScore,
          row.EffectiveCostEUR,
          row.Type,
          row.Priority,
          row.SLA,
          row.Owner,
          row.Status,
          row.Action,
          row.ResolvedAt
      ]

      for c, value in enumerate(values, start=1):
          ws_anom.cell(row=r, column=c, value=value)
  ```

  > **Salida esperada:** Se crean las hojas `Forecast` y `Anomaly Log`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Guarda el dashboard en `outputs\lab5`.

  > **Importante:** El archivo final debe contener exactamente cinco hojas.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\build_finops_dashboard.py`, agrega este bloque **al final del archivo**, debajo del bloque que crea `Anomaly Log`.
  {: .lab-note .important .compact}

  ```python
  wb.save(dashboard_xlsx)
  print(f"Dashboard creado: {dashboard_xlsx}")
  ```

  > **Salida esperada:** Se crea `outputs\lab5\novatech_finops_dashboard_v1.xlsx`.
  {: .lab-note .output .compact}

{% capture r4 %}{{ results[3] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r4 %}

{% include support-prompt.html task="tarea4" %}

---

## ✅ Tarea 5. Validar el entregable — 5 min

En esta tarea ejecutarás el script completo y verificarás la integridad de todos los artefactos generados.

### Tarea 5.1. Ejecutar y validar

- {% include step_label.html %} Ejecuta el script completo desde PowerShell.

  > **Nota:** Si aparece un `FileNotFoundError`, verifica que las prácticas anteriores hayan generado sus archivos en `data\processed`.
  {: .lab-note .info .compact}

  ```powershell
  python "$LabRoot\scripts\build_finops_dashboard.py"
  ```

  > **Salida esperada:** Se generan el dashboard Excel, la gráfica de forecast y el Anomaly Log sin errores.
  {: .lab-note .output .compact}

- {% include step_label.html %} Verifica los tres artefactos generados.

  > **Importante:** Todos los archivos deben existir y tener un tamaño mayor a 0 bytes.
  {: .lab-note .important .compact}

  ```powershell
  Get-Item `
    "$LabRoot\outputs\lab5\novatech_finops_dashboard_v1.xlsx", `
    "$LabRoot\outputs\lab5\forecast_feb_apr_2024.png", `
    "$LabRoot\outputs\lab5\anomaly_log.csv" |
    Select-Object Name, Length
  ```

  > **Salida esperada:** PowerShell muestra los tres archivos con tamaños mayores a 0 bytes.
  {: .lab-note .output .compact}

- {% include step_label.html %} Confirma que el dashboard tenga exactamente cinco hojas.

  > **Nota:** Esta verificación valida la estructura final del entregable.
  {: .lab-note .info .compact}

  ```powershell
  python -c "import openpyxl; p=r'$LabRoot\outputs\lab5\novatech_finops_dashboard_v1.xlsx'; w=openpyxl.load_workbook(p, read_only=True); print(w.sheetnames); print('Total hojas:',len(w.sheetnames))"
  ```

  > **Salida esperada:** Se muestran `Executive Summary`, `Cost & Trend`, `Budget vs Actual`, `Forecast` y `Anomaly Log`, con un total de 5 hojas.
  {: .lab-note .output .compact}

{% capture r5 %}{{ results[4] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r5 %}
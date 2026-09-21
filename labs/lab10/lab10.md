---
layout: lab
title: "Demostración 10: Caso integrador FinOps: diagnóstico, roadmap y reporte ejecutivo"
permalink: /lab10/lab10/
images_base: /labs/lab10/img
duration: "45 minutos"
objective:
  - Integrar los artefactos generados durante el curso para construir un diagnóstico FinOps basado en evidencia, un roadmap 30-60-90, un operating model con KPIs y cadencias, y un reporte ejecutivo final para NovaTech SRL.
prerequisites:
  - Equipo con Windows 11 y Visual Studio Code instalado.
  - Python 3.12 o superior disponible desde la terminal integrada de Visual Studio Code.
  - Paquetes pandas, openpyxl, pyyaml y plotly instalados.
  - Haber completado las Prácticas 3 a 9 o disponer de sus artefactos de referencia.
  - Conocimiento básico de Allocation, Budgeting, Optimization, Governance, Data Quality y FinOps Maturity.
introduction:
  - En esta práctica integradora consolidarás las evidencias producidas durante el curso para transformar datos y artefactos técnicos en decisiones. Evaluarás readiness por Capability, construirás un roadmap 30-60-90, definirás KPIs y cadencias operativas y generarás un reporte ejecutivo HTML autocontenido.
slug: lab10
lab_number: 10
final_result: >
  Al finalizar tendrás cuatro entregables de cierre: finops_diagnostic_novatech.xlsx, finops_roadmap_30_60_90.xlsx, finops_operating_model.yaml y finops_executive_report_novatech.html.
notes:
  - Todos los archivos se almacenan dentro de la carpeta finops-course del Escritorio de Windows.
  - Todos los importes monetarios se expresan en EUR.
  - El Capability Readiness Assessment es indicativo y se basa únicamente en evidencias generadas durante este curso.
  - Crawl, Walk y Run se interpretan por Capability y no como un score global único de toda la organización.
  - EstimatedMonthlySavingsEUR representa oportunidad estimada y no savings realizados.
  - No se generan datasets nuevos; esta práctica reutiliza artefactos de las prácticas anteriores.
references:
  - text: "FinOps Framework"
    url: https://www.finops.org/framework/
  - text: "FinOps Maturity Model"
    url: https://www.finops.org/framework/maturity-model/
  - text: "FinOps Governance, Policy & Risk"
    url: https://www.finops.org/framework/capabilities/governance-policy-risk/
  - text: "FinOps Budgeting"
    url: https://www.finops.org/framework/capabilities/budgeting/
  - text: "FinOps Anomaly Management"
    url: https://www.finops.org/framework/capabilities/anomaly-management/
prev: /lab9/lab9/
next: /
---

---

<!-- Aquí comienzan las instrucciones paso a paso de la práctica -->

## 🧩 Tarea 1. Consolidar evidencias y construir el diagnóstico — 10 min

En esta tarea validarás los artefactos previos del curso y construirás un diagnóstico FinOps basado únicamente en evidencia disponible.

### Tarea 1.1. Preparar el entorno y validar artefactos

- {% include step_label.html %} Abre Visual Studio Code y una terminal integrada de PowerShell.

  > **Nota:** Usa **Terminal → New Terminal** y confirma que el perfil activo sea PowerShell.
  {: .lab-note .info .compact}

  > **Salida esperada:** VS Code muestra una terminal PowerShell lista para ejecutar comandos.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la carpeta del curso y crea las carpetas de salida de la práctica.

  > **Importante:** La variable `$LabRoot` evita depender del nombre específico del usuario de Windows.
  {: .lab-note .important .compact}

  ```powershell
  $LabRoot = Join-Path ([Environment]::GetFolderPath('Desktop')) 'finops-course'

  New-Item -ItemType Directory -Force -Path `
    "$LabRoot\outputs\lab10", `
    "$LabRoot\scripts", `
    "$LabRoot\docs\lab10" | Out-Null

  Set-Location $LabRoot
  ```

  > **Salida esperada:** Existen `outputs\lab10`, `scripts` y `docs\lab10`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Valida que existan los principales artefactos generados durante el curso.

  > **Importante:** Si falta alguno, solicita al instructor el artefacto de referencia antes de continuar.
  {: .lab-note .important .compact}

  ```powershell
  $RequiredFiles = @(
    "$LabRoot\data\processed\novatech_trend_report.xlsx",
    "$LabRoot\data\processed\novatech_allocation_report.xlsx",
    "$LabRoot\outputs\lab5\novatech_finops_dashboard_v1.xlsx",
    "$LabRoot\data\processed\finops_governance_v1.xlsx",
    "$LabRoot\data\processed\finops_modern_cost_drivers.xlsx",
    "$LabRoot\data\processed\finops_optimization_backlog_enriched.xlsx"
  )

  $RequiredFiles | ForEach-Object {
    [PSCustomObject]@{
      File   = Split-Path $_ -Leaf
      Exists = Test-Path $_
    }
  } | Format-Table -AutoSize
  ```

  > **Salida esperada:** Los seis archivos muestran `Exists = True`.
  {: .lab-note .output .compact}

### Tarea 1.2. Generar diagnóstico por Capability

- {% include step_label.html %} Crea `scripts\build_finops_diagnostic.py`.

  > **Nota:** El script toma evidencia de Allocation, Budgeting, Optimization, Governance y Data para generar un readiness indicativo por Capability.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_diagnostic.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  from pathlib import Path
  import pandas as pd

  ROOT = Path.home() / "Desktop" / "finops-course"
  PROCESSED = ROOT / "data" / "processed"
  OUTPUT = ROOT / "outputs" / "lab10"

  allocation_file = PROCESSED / "novatech_allocation_report.xlsx"
  dashboard_file = ROOT / "outputs" / "lab5" / "novatech_finops_dashboard_v1.xlsx"
  governance_file = PROCESSED / "finops_governance_v1.xlsx"
  modern_file = PROCESSED / "finops_modern_cost_drivers.xlsx"
  backlog_file = PROCESSED / "finops_optimization_backlog_enriched.xlsx"

  allocation_summary = pd.read_excel(
      allocation_file,
      sheet_name="summary"
  )

  backlog = pd.read_excel(
      backlog_file,
      sheet_name="backlog"
  )

  governance = pd.read_excel(
      governance_file,
      sheet_name="governance_assessment"
  )

  allocation_metrics = dict(
      zip(
          allocation_summary["Metric"],
          allocation_summary["Value"]
      )
  )

  direct_allocation_pct = float(
      allocation_metrics.get(
          "DirectAllocationPct",
          0
      )
  )

  residual_unallocated_pct = float(
      allocation_metrics.get(
          "ResidualUnallocatedPct",
          100
      )
  )

  estimated_opportunity_eur = float(
      pd.to_numeric(
          backlog.get(
              "EstimatedMonthlySavingsEUR",
              pd.Series(dtype=float)
          ),
          errors="coerce"
      ).fillna(0).sum()
  )

  backlog_items = len(backlog)

  governance_avg = float(
      pd.to_numeric(
          governance["ReadinessPct"],
          errors="coerce"
      ).fillna(0).mean()
  )

  def maturity_from_pct(value):
      if value < 50:
          return "Crawl"
      if value < 80:
          return "Walk"
      return "Run"

  readiness = pd.DataFrame([
      {
          "Capability": "Allocation",
          "Evidence": (
              f"Direct allocation {direct_allocation_pct:.1f}% | "
              f"Residual unallocated {residual_unallocated_pct:.1f}%"
          ),
          "IndicativeReadinessPct": round(
              direct_allocation_pct,
              1
          ),
          "IndicativeMaturity": maturity_from_pct(
              direct_allocation_pct
          )
      },
      {
          "Capability": "Budgeting",
          "Evidence": (
              "Budget, actuals and forecast artifacts available "
              "from Practice 5"
          ),
          "IndicativeReadinessPct": 70.0,
          "IndicativeMaturity": "Walk"
      },
      {
          "Capability": "Optimization",
          "Evidence": (
              f"{backlog_items} backlog items | "
              f"€{estimated_opportunity_eur:,.2f} estimated opportunity"
          ),
          "IndicativeReadinessPct": 70.0,
          "IndicativeMaturity": "Walk"
      },
      {
          "Capability": "Governance",
          "Evidence": (
              f"Governance readiness average "
              f"{governance_avg:.1f}%"
          ),
          "IndicativeReadinessPct": round(
              governance_avg,
              1
          ),
          "IndicativeMaturity": maturity_from_pct(
              governance_avg
          )
      },
      {
          "Capability": "Data",
          "Evidence": (
              "Multicloud ingestion, normalization, quality "
              "and dashboard pipeline implemented in Practice 8"
          ),
          "IndicativeReadinessPct": 70.0,
          "IndicativeMaturity": "Walk"
      }
  ])

  governance_gaps = governance[
      governance["ReadinessPct"] < 100
  ].copy()

  optimization_summary = (
      backlog.groupby(
          "OptimizationCategory",
          dropna=False
      )
      .agg(
          Opportunities=("BacklogId", "count"),
          EstimatedOpportunityEUR=(
              "EstimatedMonthlySavingsEUR",
              "sum"
          )
      )
      .reset_index()
      .sort_values(
          "EstimatedOpportunityEUR",
          ascending=False
      )
  )

  executive_summary = pd.DataFrame([
      {
          "Metric": "DirectAllocationPct",
          "Value": round(direct_allocation_pct, 1)
      },
      {
          "Metric": "ResidualUnallocatedPct",
          "Value": round(residual_unallocated_pct, 1)
      },
      {
          "Metric": "OptimizationBacklogItems",
          "Value": backlog_items
      },
      {
          "Metric": "EstimatedMonthlyOpportunityEUR",
          "Value": round(estimated_opportunity_eur, 2)
      },
      {
          "Metric": "GovernanceReadinessAvgPct",
          "Value": round(governance_avg, 1)
      }
  ])

  output_file = OUTPUT / "finops_diagnostic_novatech.xlsx"

  with pd.ExcelWriter(
      output_file,
      engine="openpyxl"
  ) as writer:
      executive_summary.to_excel(
          writer,
          sheet_name="executive_summary",
          index=False
      )
      readiness.to_excel(
          writer,
          sheet_name="capability_readiness",
          index=False
      )
      governance_gaps.to_excel(
          writer,
          sheet_name="governance_gaps",
          index=False
      )
      optimization_summary.to_excel(
          writer,
          sheet_name="optimization_summary",
          index=False
      )

  print(f"Creado: {output_file}")
  print(readiness.to_string(index=False))
  ```

  > **Salida esperada:** Se crea `finops_diagnostic_novatech.xlsx` con cuatro hojas de diagnóstico.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el diagnóstico.

  ```powershell
  python "$LabRoot\scripts\build_finops_diagnostic.py"
  ```

  > **Salida esperada:** Se muestran cinco Capabilities con evidencia y maturity indicativa por Capability.
  {: .lab-note .output .compact}

{% assign results = site.data.task-results[page.slug].results %}
{% capture r1 %}{{ results[0] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r1 %}

---

## 🗺️ Tarea 2. Construir el roadmap 30-60-90 — 12 min

En esta tarea transformarás el backlog de optimización en un roadmap operativo con responsables, outcomes y métricas.

### Tarea 2.1. Clasificar iniciativas por horizonte

- {% include step_label.html %} Crea `scripts\build_finops_roadmap.py`.

  > **Nota:** El horizonte se determina por combinación de PriorityScore, EffortDays y tipo de acción.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\build_finops_roadmap.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  from pathlib import Path
  import pandas as pd

  ROOT = Path.home() / "Desktop" / "finops-course"
  PROCESSED = ROOT / "data" / "processed"
  OUTPUT = ROOT / "outputs" / "lab10"

  backlog = pd.read_excel(
      PROCESSED / "finops_optimization_backlog_enriched.xlsx",
      sheet_name="backlog"
  )

  def classify_horizon(row):
      action = str(
          row.get("ActionType", "")
      )

      effort = pd.to_numeric(
          row.get("EffortDays", 5),
          errors="coerce"
      )

      score = pd.to_numeric(
          row.get("PriorityScore", 0),
          errors="coerce"
      )

      if action in [
          "Architecture Review",
          "Value Optimization",
          "Governance"
      ] and effort >= 5:
          return 90

      if effort <= 3 and score > 0:
          return 30

      return 60

  backlog["HorizonDays"] = backlog.apply(
      classify_horizon,
      axis=1
  )

  def expected_outcome(row):
      category = str(
          row.get(
              "OptimizationCategory",
              ""
          )
      )

      mapping = {
          "Low_Utilization":
              "Validate need and remediate unused capacity",
          "Rightsizing":
              "Reduce overprovisioning after technical validation",
          "Storage":
              "Remove or tier unnecessary storage",
          "Rate_Optimization":
              "Evaluate commitment discount opportunity",
          "Governance":
              "Close ownership or policy gap",
          "Kubernetes":
              "Reduce inefficient workload or shared-cost drivers",
          "Multicloud":
              "Validate observed unit-cost variance and architecture impact",
          "AI":
              "Improve value-to-cost economics of AI workload"
      }

      return mapping.get(
          category,
          "Validate and execute approved optimization action"
      )

  backlog["ExpectedOutcome"] = backlog.apply(
      expected_outcome,
      axis=1
  )

  if "ResponsiblePersona" not in backlog.columns:
      backlog["ResponsiblePersona"] = "FinOps"

  backlog["SuccessMetric"] = backlog[
      "OptimizationCategory"
  ].map({
      "Low_Utilization": "Reviewed candidate and decision documented",
      "Rightsizing": "Post-change utilization within agreed target",
      "Storage": "Unused storage reviewed and disposition documented",
      "Rate_Optimization": "Commitment analysis approved or rejected",
      "Governance": "Ownership or policy gap closed",
      "Kubernetes": "Cost driver reviewed and action assigned",
      "Multicloud": "Comparability and architecture review completed",
      "AI": "ValueCostRatio reviewed and optimization decision documented"
  }).fillna(
      "Outcome validated and evidence recorded"
  )

  columns = [
      "BacklogId",
      "OptimizationCategory",
      "ActionType",
      "ResponsiblePersona",
      "HorizonDays",
      "ExpectedOutcome",
      "SuccessMetric",
      "EstimatedMonthlySavingsEUR",
      "EffortDays",
      "RiskLevel",
      "PriorityScore",
      "Status"
  ]

  roadmap = backlog[
      [c for c in columns if c in backlog.columns]
  ].copy()

  output_file = (
      OUTPUT
      / "finops_roadmap_30_60_90.xlsx"
  )

  with pd.ExcelWriter(
      output_file,
      engine="openpyxl"
  ) as writer:

      for horizon in [30, 60, 90]:
          horizon_df = roadmap[
              roadmap["HorizonDays"] == horizon
          ].copy()

          horizon_df = horizon_df.sort_values(
              "PriorityScore",
              ascending=False
          )

          horizon_df.to_excel(
              writer,
              sheet_name=f"{horizon}_days",
              index=False
          )

      summary = (
          roadmap.groupby(
              "HorizonDays",
              as_index=False
          )
          .agg(
              Initiatives=("BacklogId", "count"),
              EstimatedOpportunityEUR=(
                  "EstimatedMonthlySavingsEUR",
                  "sum"
              )
          )
      )

      summary.to_excel(
          writer,
          sheet_name="summary",
          index=False
      )

  print(f"Creado: {output_file}")
  print(
      roadmap[
          [
              "BacklogId",
              "OptimizationCategory",
              "HorizonDays",
              "ResponsiblePersona"
          ]
      ].to_string(index=False)
  )
  ```

  > **Salida esperada:** Cada iniciativa queda asignada a un horizonte de 30, 60 o 90 días.
  {: .lab-note .output .compact}

### Tarea 2.2. Generar y validar el roadmap

- {% include step_label.html %} Ejecuta el generador de roadmap.

  ```powershell
  python "$LabRoot\scripts\build_finops_roadmap.py"
  ```

  > **Salida esperada:** Se crea `finops_roadmap_30_60_90.xlsx`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Verifica las hojas del workbook.

  ```powershell
  python -c "import pandas as pd; p=r'$LabRoot\outputs\lab10\finops_roadmap_30_60_90.xlsx'; print(pd.ExcelFile(p).sheet_names)"
  ```

  > **Salida esperada:** Se muestran `30_days`, `60_days`, `90_days` y `summary`.
  {: .lab-note .output .compact}

{% capture r2 %}{{ results[1] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r2 %}

---

## 📏 Tarea 3. Definir KPIs y cadencias operativas — 8 min

En esta tarea construirás el operating model mínimo para dar continuidad al trabajo FinOps después del curso.

### Tarea 3.1. Crear KPIs

- {% include step_label.html %} Crea `docs\lab10\finops_operating_model.yaml`.

  > **Importante:** Los thresholds son objetivos internos de NovaTech y no valores universales del FinOps Framework.
  {: .lab-note .important .compact}

  > **Ubicación del contenido:** En `docs\lab10\finops_operating_model.yaml`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```yaml
  organization: "NovaTech SRL"
  currency: "EUR"
  version: "1.0"

  kpis:
    - id: KPI-001
      name: "Direct Allocation %"
      description: "Porcentaje del costo asignado directamente a dimensiones organizacionales."
      source: "novatech_allocation_report.xlsx"
      target_novatech: ">= 90%"
      owner: "Finance"
      cadence: "weekly"

    - id: KPI-002
      name: "Residual Unallocated %"
      description: "Porcentaje de costo que permanece sin asignación después del proceso de Allocation."
      source: "novatech_allocation_report.xlsx"
      target_novatech: "<= 10%"
      owner: "Finance"
      cadence: "weekly"

    - id: KPI-003
      name: "Budget Variance %"
      description: "Diferencia porcentual entre gasto observado y presupuesto."
      source: "novatech_finops_dashboard_v1.xlsx"
      target_novatech: "<= 10%"
      owner: "Finance"
      cadence: "weekly"

    - id: KPI-004
      name: "Anomaly Triage Compliance %"
      description: "Porcentaje de anomalías atendidas dentro del SLA definido por prioridad."
      source: "Anomaly Log"
      target_novatech: ">= 90%"
      owner: "Platform"
      cadence: "weekly"

    - id: KPI-005
      name: "Optimization Backlog Progress %"
      description: "Porcentaje de oportunidades revisadas o ejecutadas respecto del backlog total."
      source: "finops_optimization_backlog_enriched.xlsx"
      target_novatech: ">= 80%"
      owner: "Engineering + FinOps"
      cadence: "weekly"

  cadences:
    - id: CAD-W01
      name: "FinOps Weekly Review"
      frequency: "weekly"
      duration_minutes: 30
      owner: "FinOps"
      participants:
        - Finance
        - Engineering
        - Platform
      agenda:
        - "KPIs y budget variance"
        - "Bloqueos y decisiones"
        - "Acciones de la semana"
      output: "Weekly action log"

    - id: CAD-W02
      name: "Anomaly Triage"
      frequency: "weekly"
      duration_minutes: 20
      owner: "Platform"
      participants:
        - Platform
        - Engineering
        - Finance
      agenda:
        - "Revisar anomalías abiertas"
        - "Validar impacto"
        - "Asignar owner y SLA"
      output: "Anomaly log actualizado"

    - id: CAD-W03
      name: "Optimization Backlog Review"
      frequency: "weekly"
      duration_minutes: 30
      owner: "Engineering"
      participants:
        - Engineering
        - Platform
        - FinOps
      agenda:
        - "Nuevas oportunidades"
        - "Repriorización"
        - "Estado de iniciativas"
      output: "Backlog actualizado"

    - id: CAD-M01
      name: "Cost & Forecast Review"
      frequency: "monthly"
      duration_minutes: 45
      owner: "Finance"
      participants:
        - Finance
        - Product
        - Leadership
        - FinOps
      agenda:
        - "Actual vs Budget"
        - "Forecast"
        - "Top cost drivers"
        - "Decisiones requeridas"
      output: "Monthly cost review"

    - id: CAD-M02
      name: "Optimization Planning"
      frequency: "monthly"
      duration_minutes: 45
      owner: "FinOps"
      participants:
        - Engineering
        - Platform
        - Finance
      agenda:
        - "Revisión del mes anterior"
        - "Selección de iniciativas"
        - "Asignación de owners"
      output: "Monthly optimization plan"

  decision_rules:
    - "No presentar EstimatedMonthlySavingsEUR como savings realizados."
    - "Multicloud price variance requiere validación técnica antes de estimar savings."
    - "Toda excepción de Governance debe tener owner y fecha de revisión."
    - "Toda iniciativa del roadmap debe tener ExpectedOutcome y SuccessMetric."
  ```

  > **Salida esperada:** Se define un operating model con cinco KPIs, cinco cadencias y reglas de decisión.
  {: .lab-note .output .compact}

### Tarea 3.2. Validar el operating model

- {% include step_label.html %} Valida el YAML desde PowerShell.

  ```powershell
  python -c "import yaml, pathlib; p=pathlib.Path(r'$LabRoot\docs\lab10\finops_operating_model.yaml'); d=yaml.safe_load(p.read_text(encoding='utf-8')); print('KPIs:', len(d['kpis'])); print('Cadences:', len(d['cadences'])); print('Decision rules:', len(d['decision_rules']))"
  ```

  > **Salida esperada:** Se muestran 5 KPIs, 5 cadencias y 4 reglas de decisión.
  {: .lab-note .output .compact}

{% capture r3 %}{{ results[2] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r3 %}

---

## 📊 Tarea 4. Generar el reporte ejecutivo — 12 min

En esta tarea generarás un reporte HTML autocontenido utilizando únicamente datos y artefactos construidos durante el curso.

### Tarea 4.1. Crear el generador HTML

- {% include step_label.html %} Crea `scripts\generate_executive_report.py`.

  > **Nota:** El reporte evita datos ficticios y utiliza directamente los outputs de las tareas anteriores.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\generate_executive_report.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  {%raw%}
  ```python
  from pathlib import Path
  import json
  import pandas as pd
  import plotly.express as px
  import yaml

  ROOT = Path.home() / "Desktop" / "finops-course"
  OUTPUT = ROOT / "outputs" / "lab10"
  DOCS = ROOT / "docs" / "lab10"

  diagnostic_file = (
      OUTPUT
      / "finops_diagnostic_novatech.xlsx"
  )

  roadmap_file = (
      OUTPUT
      / "finops_roadmap_30_60_90.xlsx"
  )

  operating_model_file = (
      DOCS
      / "finops_operating_model.yaml"
  )

  readiness = pd.read_excel(
      diagnostic_file,
      sheet_name="capability_readiness"
  )

  optimization = pd.read_excel(
      diagnostic_file,
      sheet_name="optimization_summary"
  )

  roadmap_summary = pd.read_excel(
      roadmap_file,
      sheet_name="summary"
  )

  with open(
      operating_model_file,
      "r",
      encoding="utf-8"
  ) as f:
      operating_model = yaml.safe_load(f)

  fig_readiness = px.bar(
      readiness,
      x="Capability",
      y="IndicativeReadinessPct",
      color="IndicativeMaturity",
      title="Capability Readiness - NovaTech SRL",
      range_y=[0, 100]
  )

  fig_optimization = px.bar(
      optimization,
      x="OptimizationCategory",
      y="EstimatedOpportunityEUR",
      title="Estimated Opportunity by Category",
      labels={
          "EstimatedOpportunityEUR":
              "Estimated Opportunity (EUR)"
      }
  )

  fig_roadmap = px.bar(
      roadmap_summary,
      x="HorizonDays",
      y="Initiatives",
      title="Roadmap Initiatives by Horizon",
      text="Initiatives"
  )

  figures = {
      "readiness": fig_readiness.to_html(
          full_html=False,
          include_plotlyjs="inline"
      ),
      "optimization": fig_optimization.to_html(
          full_html=False,
          include_plotlyjs=False
      ),
      "roadmap": fig_roadmap.to_html(
          full_html=False,
          include_plotlyjs=False
      )
  }

  kpi_rows = "".join(
      f"""
      <tr>
        <td>{kpi['name']}</td>
        <td>{kpi['target_novatech']}</td>
        <td>{kpi['owner']}</td>
        <td>{kpi['cadence']}</td>
      </tr>
      """
      for kpi in operating_model["kpis"]
  )

  html = f"""
  <!DOCTYPE html>
  <html lang="es">
  <head>
    <meta charset="utf-8">
    <title>FinOps Executive Report - NovaTech SRL</title>
    <style>
      body {{
        font-family: Arial, sans-serif;
        margin: 40px;
        color: #222;
        background: #fff;
      }}
      h1, h2 {{
        color: #1f2937;
      }}
      .card {{
        border: 1px solid #ddd;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 24px;
      }}
      table {{
        width: 100%;
        border-collapse: collapse;
      }}
      th, td {{
        border-bottom: 1px solid #ddd;
        text-align: left;
        padding: 10px;
      }}
      th {{
        background: #f3f4f6;
      }}
      .note {{
        font-size: 0.92rem;
        color: #555;
      }}
    </style>
  </head>
  <body>

    <h1>FinOps Executive Report - NovaTech SRL</h1>

    <p>
      Este reporte consolida las evidencias generadas durante el curso
      para apoyar decisiones de Allocation, Budgeting, Optimization,
      Governance y Data.
    </p>

    <div class="card">
      <h2>1. Capability Readiness</h2>
      {figures['readiness']}
      <p class="note">
        El assessment es indicativo y se aplica por Capability.
        No representa un score global certificado de madurez.
      </p>
    </div>

    <div class="card">
      <h2>2. Optimization Opportunity</h2>
      {figures['optimization']}
      <p class="note">
        EstimatedOpportunityEUR representa oportunidad estimada,
        no savings realizados.
      </p>
    </div>

    <div class="card">
      <h2>3. Roadmap 30-60-90</h2>
      {figures['roadmap']}
    </div>

    <div class="card">
      <h2>4. KPIs Operativos</h2>
      <table>
        <thead>
          <tr>
            <th>KPI</th>
            <th>NovaTech Target</th>
            <th>Owner</th>
            <th>Cadence</th>
          </tr>
        </thead>
        <tbody>
          {kpi_rows}
        </tbody>
      </table>
    </div>

    <div class="card">
      <h2>5. Próximos pasos</h2>
      <ul>
        <li>Ejecutar las iniciativas del horizonte de 30 días.</li>
        <li>Medir resultados reales antes de declarar savings.</li>
        <li>Revisar gaps de Governance y Allocation semanalmente.</li>
        <li>Actualizar el roadmap al cierre de cada ciclo mensual.</li>
      </ul>
    </div>

  </body>
  </html>
  """

  output_file = (
      OUTPUT
      / "finops_executive_report_novatech.html"
  )

  output_file.write_text(
      html,
      encoding="utf-8"
  )

  print(f"Creado: {output_file}")
  ```
  {%endraw%}

  > **Salida esperada:** El script queda preparado para generar el reporte ejecutivo final.
  {: .lab-note .output .compact}

### Tarea 4.2. Generar el reporte

- {% include step_label.html %} Ejecuta el generador HTML.

  ```powershell
  python "$LabRoot\scripts\generate_executive_report.py"
  ```

  > **Salida esperada:** Se crea `finops_executive_report_novatech.html`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Abre el reporte en el navegador.

  ```powershell
  Start-Process "$LabRoot\outputs\lab10\finops_executive_report_novatech.html"
  ```

  > **Salida esperada:** El navegador muestra Capability Readiness, oportunidades, roadmap y KPIs.
  {: .lab-note .output .compact}

{% capture r4 %}{{ results[3] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r4 %}

---

## ✅ Tarea 5. Validar los entregables finales — 3 min

En esta tarea comprobarás que los cuatro artefactos de cierre existen y contienen la estructura esperada.

### Tarea 5.1. Ejecutar validación integral

- {% include step_label.html %} Ejecuta la validación final.

  > **Importante:** Los cuatro archivos deben existir y tener un tamaño mayor a 0 bytes.
  {: .lab-note .important .compact}

  ```powershell
  $Artifacts = @(
    "$LabRoot\outputs\lab10\finops_diagnostic_novatech.xlsx",
    "$LabRoot\outputs\lab10\finops_roadmap_30_60_90.xlsx",
    "$LabRoot\docs\lab10\finops_operating_model.yaml",
    "$LabRoot\outputs\lab10\finops_executive_report_novatech.html"
  )

  $Artifacts | ForEach-Object {
    Get-Item $_ | Select-Object Name, Length
  }

  python -c "import pandas as pd; p=r'$LabRoot\outputs\lab10\finops_diagnostic_novatech.xlsx'; print('Diagnostic:', pd.ExcelFile(p).sheet_names)"

  python -c "import pandas as pd; p=r'$LabRoot\outputs\lab10\finops_roadmap_30_60_90.xlsx'; print('Roadmap:', pd.ExcelFile(p).sheet_names)"
  ```

  > **Salida esperada:** Los cuatro artefactos existen; el diagnóstico tiene cuatro hojas y el roadmap contiene 30, 60, 90 días y summary.
  {: .lab-note .output .compact}

### Tarea 5.2. Confirmar el resultado integrador

- {% include step_label.html %} Revisa que los entregables finales representen las cuatro capas del caso integrador.

  > **Salida esperada:** El diagnóstico describe el estado actual, el roadmap define acciones futuras, el operating model establece la cadencia de gestión y el HTML comunica los resultados a stakeholders.
  {: .lab-note .output .compact}

{% capture r5 %}{{ results[4] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r5 %}
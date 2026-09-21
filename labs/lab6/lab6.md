---
layout: lab
title: "Demostración 6: Backlog de optimización y priorización por impacto"
permalink: /lab6/lab6/
images_base: /labs/lab6/img
duration: "30 minutos"
objective:
  - Construir un backlog de oportunidades de optimización FinOps a partir de métricas de utilización, costo, ownership, storage y pricing, priorizando iniciativas por ahorro estimado, esfuerzo y riesgo.
prerequisites:
  - Equipo con Windows 11 y Visual Studio Code instalado.
  - Python 3.12 o superior disponible desde la terminal integrada de Visual Studio Code.
  - Paquetes pandas, numpy, matplotlib y openpyxl instalados.
  - Conocimiento básico de Usage Optimization, Rate Optimization, rightsizing y shared responsibility en FinOps.
introduction:
  - En esta demostración analizarás 155 recursos tecnológicos de NovaTech SRL para identificar oportunidades de Low Utilization, Rightsizing, Storage, Rate Optimization y Governance. Después construirás un backlog priorizado y generarás visualizaciones y un reporte Excel para apoyar decisiones de optimización.
slug: lab6
lab_number: 6
final_result: >
  Al finalizar tendrás un backlog de optimización en Excel con resumen ejecutivo, detalle de oportunidades, categorías y quick wins, además de dos visualizaciones PNG para apoyar la priorización.
notes:
  - Todos los archivos se almacenan dentro de la carpeta finops-course del Escritorio de Windows.
  - Todos los importes monetarios se expresan en EUR.
  - Las reglas de detección son simplificaciones didácticas y deben interpretarse como candidatos a revisión, no como acciones automáticas.
  - Los factores de ahorro, esfuerzo, riesgo y Priority Score son supuestos internos de NovaTech para esta demostración.
  - Rate Optimization identifica candidatos para evaluar commitments; no implica que deban comprarse automáticamente.
references:
  - text: "FinOps Usage Optimization"
    url: https://www.finops.org/framework/capabilities/usage-optimization/
  - text: "FinOps Rate Optimization"
    url: https://www.finops.org/framework/capabilities/rate-optimization/
  - text: "FinOps Optimization"
    url: https://www.finops.org/framework/
  - text: "pandas Documentation"
    url: https://pandas.pydata.org/docs/
prev: /lab5/lab5/
next: /lab7/lab7/
---

---

<!-- Aquí comienzan las instrucciones paso a paso de la práctica -->

## 🔎 Tarea 1. Preparar y explorar los recursos — 5 min

En esta tarea descargarás el dataset de optimización, validarás su estructura y revisarás las métricas disponibles para identificar candidatos de eficiencia.

### Tarea 1.1. Preparar el entorno y descargar el dataset

- {% include step_label.html %} Abre Visual Studio Code y una terminal integrada de PowerShell.

  > **Nota:** Usa **Terminal → New Terminal** y confirma que el perfil activo sea PowerShell.
  {: .lab-note .info .compact}

  > **Salida esperada:** VS Code muestra una terminal PowerShell lista para ejecutar comandos.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la carpeta del laboratorio y crea los directorios necesarios.

  > **Importante:** La variable `$LabRoot` permite trabajar sin depender del nombre específico del usuario de Windows.
  {: .lab-note .important .compact}

  ```powershell
  $LabRoot = Join-Path ([Environment]::GetFolderPath('Desktop')) 'finops-course'
  New-Item -ItemType Directory -Force -Path `
    "$LabRoot\data\raw", `
    "$LabRoot\data\processed", `
    "$LabRoot\outputs\lab6", `
    "$LabRoot\scripts" | Out-Null
  Set-Location $LabRoot
  ```

  > **Salida esperada:** Existen `data\raw`, `data\processed`, `outputs\lab6` y `scripts`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la URL del dataset y descárgalo en `data\raw`.

  ```powershell
  $DatasetUrl = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_optimization_resources_eur.csv"

  Invoke-WebRequest `
    -Uri $DatasetUrl `
    -OutFile "$LabRoot\data\raw\novatech_optimization_resources_eur.csv"
  ```

  > **Salida esperada:** Se crea `data\raw\novatech_optimization_resources_eur.csv`.
  {: .lab-note .output .compact}

### Tarea 1.2. Validar estructura y costo total

- {% include step_label.html %} Verifica cantidad de recursos, costo mensual y columnas clave.

  > **Nota:** El dataset contiene señales de utilización, ownership, attachment, antigüedad y pricing.
  {: .lab-note .info .compact}

  ```powershell
  python -c "import pandas as pd; p=r'$LabRoot\data\raw\novatech_optimization_resources_eur.csv'; df=pd.read_csv(p); print('Recursos:',len(df)); print('Costo mensual EUR:',round(df['MonthlyEffectiveCostEUR'].sum(),2)); print('Tipos:',sorted(df['ResourceType'].unique())); print('Columnas:',df.columns.tolist())"
  ```

  > **Salida esperada:** Se muestran 155 recursos, costo mensual cercano a €54,846 y los tipos VM, Database, Storage_Volume, Elastic_IP, Load_Balancer, Snapshot y Container.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea `scripts\optimization_backlog.py` y carga el dataset.

  > **Nota:** A partir de este archivo construirás todo el análisis de optimización.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  from pathlib import Path
  import pandas as pd
  import numpy as np
  import matplotlib.pyplot as plt

  ROOT = Path.home() / "Desktop" / "finops-course"
  RAW = ROOT / "data" / "raw" / "novatech_optimization_resources_eur.csv"
  PROCESSED = ROOT / "data" / "processed"
  OUTPUTS = ROOT / "outputs" / "lab6"

  PROCESSED.mkdir(parents=True, exist_ok=True)
  OUTPUTS.mkdir(parents=True, exist_ok=True)

  df = pd.read_csv(RAW)

  print(f"Recursos cargados: {len(df)}")
  print(f"Costo mensual total: €{df['MonthlyEffectiveCostEUR'].sum():,.2f}")
  ```

  > **Salida esperada:** El script carga correctamente los 155 recursos.
  {: .lab-note .output .compact}

{% assign results = site.data.task-results[page.slug].results %}
{% capture r1 %}{{ results[0] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r1 %}

---

## 🧮 Tarea 2. Detectar oportunidades de optimización — 9 min

En esta tarea aplicarás reglas didácticas para identificar candidatos de Low Utilization, Rightsizing, Storage, Rate Optimization y Governance.

### Tarea 2.1. Detectar oportunidades de Usage Optimization

- {% include step_label.html %} Agrega la detección de candidatos por baja utilización y rightsizing.

  > **Importante:** Una señal de baja utilización no implica apagar automáticamente un recurso; primero debe validarse su función y patrón de carga.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo de** `print(f"Costo mensual total: €{df['MonthlyEffectiveCostEUR'].sum():,.2f}")`.
  {: .lab-note .important .compact}

  ```python
  low_utilization = df[
      (df["ResourceType"].isin(["VM", "Container"]))
      & (df["CpuAvgPct"].notna())
      & (df["CpuAvgPct"] < 5)
  ].copy()

  low_utilization["OptimizationCategory"] = "Low_Utilization"
  low_utilization["RecommendedAction"] = "Revisar necesidad y posible apagado"

  rightsizing = df[
      (df["ResourceType"].isin(["VM", "Container"]))
      & (df["CpuAvgPct"].notna())
      & (df["MemoryAvgPct"].notna())
      & (df["CpuAvgPct"] >= 5)
      & (df["CpuAvgPct"] < 30)
      & (df["MemoryAvgPct"] < 40)
  ].copy()

  rightsizing["OptimizationCategory"] = "Rightsizing"
  rightsizing["RecommendedAction"] = "Evaluar reducción de tamaño"
  ```

  > **Salida esperada:** Se identifican candidatos de baja utilización y rightsizing sin modificar recursos.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega la detección de candidatos de Storage.

  > **Nota:** La eliminación o tiering debe validarse contra políticas de retención y requisitos técnicos.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo de** `rightsizing["RecommendedAction"] = "Evaluar reducción de tamaño"`.
  {: .lab-note .important .compact}

  ```python
  storage = df[
      (
          (df["ResourceType"] == "Storage_Volume")
          & (df["IsAttached"] == False)
      )
      |
      (
          (df["ResourceType"] == "Snapshot")
          & (df["DaysSinceCreation"] > 90)
          & (df["LastAccessDaysAgo"] > 60)
      )
  ].copy()

  storage["OptimizationCategory"] = "Storage"
  storage["RecommendedAction"] = "Evaluar eliminación o tiering"
  ```

  > **Salida esperada:** Se detectan volúmenes sin adjuntar y snapshots antiguos con baja actividad.
  {: .lab-note .output .compact}

### Tarea 2.2. Detectar Rate Optimization y Governance

- {% include step_label.html %} Agrega la detección de candidatos de Rate Optimization.

  > **Importante:** El resultado identifica candidatos para análisis de commitments; no implica una compra automática.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo de** `storage["RecommendedAction"] = "Evaluar eliminación o tiering"`.
  {: .lab-note .important .compact}

  ```python
  rate = df[
      (df["Environment"] == "prod")
      & (df["PricingModel"] == "OnDemand")
      & (df["CommitmentEligible"] == True)
      & (df["MonthlyEffectiveCostEUR"] >= 250)
  ].copy()

  rate["OptimizationCategory"] = "Rate_Optimization"
  rate["RecommendedAction"] = "Evaluar commitment discount"
  ```

  > **Salida esperada:** Se identifican recursos productivos OnDemand elegibles para evaluar optimización de tarifa.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega la detección de gaps de Governance.

  > **Nota:** La ausencia de owner es un problema de gobernanza, no evidencia de que el recurso deba eliminarse.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo de** `rate["RecommendedAction"] = "Evaluar commitment discount"`.
  {: .lab-note .important .compact}

  ```python
  governance = df[
      df["HasOwner"] == False
  ].copy()

  governance["OptimizationCategory"] = "Governance"
  governance["RecommendedAction"] = "Identificar owner y validar necesidad"
  ```

  > **Salida esperada:** Se identifican recursos sin owner para seguimiento organizacional.
  {: .lab-note .output .compact}

- {% include step_label.html %} Excluye duplicados entre categorías y consolida oportunidades.

  > **Importante:** Cada recurso se asigna a una categoría principal para evitar doble conteo del ahorro estimado.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo de** `governance["RecommendedAction"] = "Identificar owner y validar necesidad"`.
  {: .lab-note .important .compact}

  ```python
  used = set()

  categories = []

  for frame in [
      low_utilization,
      rightsizing,
      storage,
      rate,
      governance
  ]:
      frame = frame[
          ~frame["ResourceId"].isin(used)
      ].copy()

      used.update(frame["ResourceId"])
      categories.append(frame)

  opportunities = pd.concat(categories, ignore_index=True)

  print(
      opportunities["OptimizationCategory"]
      .value_counts()
      .to_string()
  )
  ```

  > **Salida esperada:** Cada recurso aparece como máximo una vez en el conjunto de oportunidades.
  {: .lab-note .output .compact}

{% capture r2 %}{{ results[1] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r2 %}

---

## 📋 Tarea 3. Construir y priorizar el backlog — 8 min

En esta tarea estimarás ahorro potencial, esfuerzo y riesgo, y calcularás un Priority Score interno para ordenar las oportunidades.

### Tarea 3.1. Estimar ahorro, esfuerzo y riesgo

- {% include step_label.html %} Agrega los supuestos de ahorro estimado por categoría.

  > **Importante:** Estos porcentajes son supuestos didácticos de NovaTech y no representan ahorros garantizados.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo del bloque que imprime** `opportunities["OptimizationCategory"].value_counts()`.
  {: .lab-note .important .compact}

  ```python
  savings_factor = {
      "Low_Utilization": 0.70,
      "Rightsizing": 0.30,
      "Storage": 0.60,
      "Rate_Optimization": 0.20,
      "Governance": 0.00
  }

  effort_days = {
      "Low_Utilization": 1,
      "Rightsizing": 3,
      "Storage": 2,
      "Rate_Optimization": 5,
      "Governance": 2
  }

  risk_level = {
      "Low_Utilization": "Medium",
      "Rightsizing": "Medium",
      "Storage": "Low",
      "Rate_Optimization": "Low",
      "Governance": "Low"
  }

  action_type = {
      "Low_Utilization": "Efficiency",
      "Rightsizing": "Efficiency",
      "Storage": "Elimination",
      "Rate_Optimization": "Rate",
      "Governance": "Governance"
  }

  opportunities["SavingsFactor"] = (
      opportunities["OptimizationCategory"]
      .map(savings_factor)
  )

  opportunities["EstimatedMonthlySavingsEUR"] = (
      opportunities["MonthlyEffectiveCostEUR"]
      * opportunities["SavingsFactor"]
  ).round(2)

  opportunities["EffortDays"] = (
      opportunities["OptimizationCategory"]
      .map(effort_days)
  )

  opportunities["RiskLevel"] = (
      opportunities["OptimizationCategory"]
      .map(risk_level)
  )

  opportunities["ActionType"] = (
      opportunities["OptimizationCategory"]
      .map(action_type)
  )
  ```

  > **Salida esperada:** Cada oportunidad contiene ahorro estimado, esfuerzo, riesgo y tipo de acción.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega el cálculo del NovaTech Priority Score.

  > **Nota:** El score permite ordenar oportunidades de forma consistente considerando ahorro, esfuerzo y riesgo.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo de** `opportunities["ActionType"] = (...)` y después de asignar `ActionType`.
  {: .lab-note .important .compact}

  ```python
  risk_factor = {
      "Low": 1.0,
      "Medium": 0.7,
      "High": 0.4
  }

  opportunities["RiskFactor"] = (
      opportunities["RiskLevel"]
      .map(risk_factor)
  )

  opportunities["PriorityScore"] = (
      opportunities["EstimatedMonthlySavingsEUR"]
      / opportunities["EffortDays"]
      * opportunities["RiskFactor"]
  ).round(2)

  opportunities["Status"] = "Pending Review"

  backlog = opportunities.sort_values(
      "PriorityScore",
      ascending=False
  ).reset_index(drop=True)

  backlog.insert(0, "BacklogId", range(1, len(backlog) + 1))
  ```

  > **Salida esperada:** El backlog queda ordenado de mayor a menor Priority Score.
  {: .lab-note .output .compact}

### Tarea 3.2. Revisar quick wins y resumen

- {% include step_label.html %} Muestra los diez candidatos con mayor prioridad.

  > **Nota:** Los quick wins son candidatos de alta prioridad para revisión, no acciones que deban ejecutarse automáticamente.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo de** `backlog.insert(0, "BacklogId", range(1, len(backlog) + 1))`.
  {: .lab-note .important .compact}

  ```python
  print(
      backlog[
          [
              "BacklogId",
              "ResourceId",
              "OptimizationCategory",
              "EstimatedMonthlySavingsEUR",
              "EffortDays",
              "RiskLevel",
              "PriorityScore"
          ]
      ]
      .head(10)
      .to_string(index=False)
  )
  ```

  > **Salida esperada:** Se muestran las diez oportunidades con mayor score de priorización.
  {: .lab-note .output .compact}

- {% include step_label.html %} Genera un resumen por categoría.

  > **Nota:** Governance puede mostrar ahorro estimado 0 porque su objetivo principal es mejorar accountability y calidad de metadata.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo del bloque que imprime los diez candidatos con mayor prioridad**.
  {: .lab-note .important .compact}

  ```python
  category_summary = (
      backlog.groupby("OptimizationCategory", as_index=False)
      .agg(
          Opportunities=("ResourceId", "count"),
          AffectedCostEUR=("MonthlyEffectiveCostEUR", "sum"),
          EstimatedMonthlySavingsEUR=("EstimatedMonthlySavingsEUR", "sum"),
          AvgPriorityScore=("PriorityScore", "mean")
      )
      .sort_values("EstimatedMonthlySavingsEUR", ascending=False)
  )

  print(category_summary.round(2).to_string(index=False))
  ```

  > **Salida esperada:** Se obtiene un resumen de oportunidades, costo afectado, ahorro estimado y score promedio por categoría.
  {: .lab-note .output .compact}

{% capture r3 %}{{ results[2] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r3 %}

---

## 📊 Tarea 4. Visualizar y exportar el backlog — 8 min

En esta tarea generarás dos visualizaciones y exportarás un reporte Excel con resumen, detalle, categorías y quick wins.

### Tarea 4.1. Generar visualizaciones de priorización

- {% include step_label.html %} Crea la gráfica de impacto vs esfuerzo.

  > **Nota:** El tamaño de cada burbuja representa el Priority Score de la oportunidad.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo de** `print(category_summary.round(2).to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  plot_data = backlog[
      backlog["EstimatedMonthlySavingsEUR"] > 0
  ].copy()

  sizes = (
      (plot_data["PriorityScore"] - plot_data["PriorityScore"].min())
      / (
          plot_data["PriorityScore"].max()
          - plot_data["PriorityScore"].min()
          + 1e-9
      )
      * 250
      + 40
  )

  plt.figure(figsize=(10, 6))

  for category, group in plot_data.groupby("OptimizationCategory"):
      idx = group.index
      plt.scatter(
          group["EffortDays"],
          group["EstimatedMonthlySavingsEUR"],
          s=sizes.loc[idx],
          alpha=0.65,
          label=category
      )

  plt.xlabel("Esfuerzo estimado (días)")
  plt.ylabel("Ahorro mensual estimado (EUR)")
  plt.legend()
  plt.grid(alpha=0.3)
  plt.tight_layout()

  priority_png = OUTPUTS / "optimization_priority.png"
  plt.savefig(priority_png, dpi=150, bbox_inches="tight")
  plt.close()

  print(f"Gráfica creada: {priority_png}")
  ```

  > **Salida esperada:** Se crea `outputs\lab6\optimization_priority.png`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea la gráfica de ahorro estimado por categoría.

  > **Nota:** Esta visualización facilita identificar qué categorías concentran mayor ahorro potencial.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo de** `print(f"Gráfica creada: {priority_png}")`.
  {: .lab-note .important .compact}

  ```python
  chart_data = category_summary.sort_values(
      "EstimatedMonthlySavingsEUR"
  )

  plt.figure(figsize=(10, 6))
  plt.barh(
      chart_data["OptimizationCategory"],
      chart_data["EstimatedMonthlySavingsEUR"]
  )

  plt.xlabel("Ahorro mensual estimado (EUR)")
  plt.tight_layout()

  savings_png = OUTPUTS / "savings_by_category.png"
  plt.savefig(savings_png, dpi=150, bbox_inches="tight")
  plt.close()

  print(f"Gráfica creada: {savings_png}")
  ```

  > **Salida esperada:** Se crea `outputs\lab6\savings_by_category.png`.
  {: .lab-note .output .compact}

### Tarea 4.2. Exportar y validar el reporte

- {% include step_label.html %} Exporta `finops_optimization_backlog.xlsx` con cuatro hojas.

  > **Nota:** El reporte conserva un resumen ejecutivo, backlog completo, resumen por categoría y top quick wins.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\optimization_backlog.py`, agrega este bloque **debajo de** `print(f"Gráfica creada: {savings_png}")`.
  {: .lab-note .important .compact}

  ```python
  report_xlsx = PROCESSED / "finops_optimization_backlog.xlsx"

  summary = pd.DataFrame({
      "Metric": [
          "TotalResources",
          "TotalMonthlyEffectiveCostEUR",
          "OptimizationOpportunities",
          "EstimatedMonthlySavingsEUR",
          "EstimatedAnnualSavingsEUR"
      ],
      "Value": [
          len(df),
          round(df["MonthlyEffectiveCostEUR"].sum(), 2),
          len(backlog),
          round(backlog["EstimatedMonthlySavingsEUR"].sum(), 2),
          round(backlog["EstimatedMonthlySavingsEUR"].sum() * 12, 2)
      ]
  })

  quick_wins = backlog.head(20).copy()

  with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
      summary.to_excel(writer, sheet_name="summary", index=False)
      backlog.to_excel(writer, sheet_name="backlog", index=False)
      category_summary.to_excel(writer, sheet_name="category_summary", index=False)
      quick_wins.to_excel(writer, sheet_name="quick_wins", index=False)

  print(f"Reporte creado: {report_xlsx}")
  ```

  > **Salida esperada:** Se crea `data\processed\finops_optimization_backlog.xlsx` con cuatro hojas.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el script completo y valida los tres artefactos.

  > **Importante:** El Excel y las dos imágenes deben existir y tener tamaño mayor a 0 bytes.
  {: .lab-note .important .compact}

  ```powershell
  python "$LabRoot\scripts\optimization_backlog.py"
  ```
  ```powershell
  Get-Item `
    "$LabRoot\data\processed\finops_optimization_backlog.xlsx", `
    "$LabRoot\outputs\lab6\optimization_priority.png", `
    "$LabRoot\outputs\lab6\savings_by_category.png" |
    Select-Object Name, Length
  ```

  > **Salida esperada:** PowerShell muestra los tres archivos con tamaños mayores a 0 bytes.
  {: .lab-note .output .compact}

{% capture r4 %}{{ results[3] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r4 %}
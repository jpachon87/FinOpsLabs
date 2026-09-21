---
layout: lab
title: "Demostración 9: Identificar drivers de costo en una solución moderna"
permalink: /lab9/lab9/
images_base: /labs/lab9/img
duration: "15 minutos"
objective:
  - Identificar drivers de costo en Kubernetes, servicios multicloud y cargas de trabajo de inteligencia artificial, cuantificando su impacto y convirtiendo los hallazgos prioritarios en nuevas oportunidades para el backlog FinOps.
prerequisites:
  - Equipo con Windows 11 y Visual Studio Code instalado.
  - Python 3.12 o superior disponible desde la terminal integrada de Visual Studio Code.
  - Paquetes pandas, numpy, matplotlib y openpyxl instalados.
  - Haber completado la Práctica 6 o disponer de `finops_optimization_backlog.xlsx`.
  - Conocimiento básico de Kubernetes, normalización multicloud y métricas de costo para workloads de AI.
introduction:
  - En esta demostración analizarás tres grupos de drivers de costo modernos para NovaTech SRL. Compararás métodos de asignación de shared costs en Kubernetes, normalizarás costos multicloud a EUR, evaluarás economics de workloads de AI y enriquecerás el backlog existente con oportunidades derivadas de los hallazgos.
slug: lab9
lab_number: 9
final_result: >
  Al finalizar tendrás un archivo finops_modern_cost_drivers.xlsx con análisis de Kubernetes, Multicloud, AI y nuevas oportunidades de optimización, además de un backlog enriquecido compatible con la Práctica 6.
notes:
  - Todos los archivos se almacenan dentro de la carpeta finops-course del Escritorio de Windows.
  - EUR es la moneda base del laboratorio.
  - Las tasas FX son fijas y didácticas para garantizar reproducibilidad; no representan tipos de cambio actuales.
  - Los métodos de asignación de shared costs representan políticas posibles de NovaTech y no una regla universal.
  - La comparación multicloud identifica variaciones observadas de precio y no implica que el proveedor con menor costo unitario sea automáticamente la mejor opción.
  - AttributedBusinessValueEUR es un supuesto didáctico para comparar valor y costo de workloads de AI.
references:
  - text: "OpenCost Specification"
    url: https://opencost.io/docs/specification/
  - text: "FinOps AI"
    url: https://www.finops.org/framework/technology-categories/ai/
  - text: "FOCUS Specification"
    url: https://focus.finops.org/
  - text: "FinOps Framework"
    url: https://www.finops.org/framework/
prev: /lab8/lab8/
next: /lab10/lab10/
---

---

<!-- Aquí comienzan las instrucciones paso a paso de la práctica -->

## ☸️ Tarea 1. Identificar drivers de costo Kubernetes — 5 min

En esta tarea cargarás el dataset Kubernetes, identificarás los namespaces con mayor costo y compararás tres métodos de asignación de costos compartidos.

### Tarea 1.1. Preparar el entorno y descargar los archivos

- {% include step_label.html %} Abre Visual Studio Code y una terminal integrada de PowerShell.

  > **Nota:** Usa **Terminal → New Terminal** y confirma que el perfil activo sea PowerShell.
  {: .lab-note .info .compact}

  > **Salida esperada:** VS Code muestra una terminal PowerShell lista para ejecutar comandos.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la carpeta del laboratorio y crea los directorios necesarios.

  > **Importante:** La variable `$LabRoot` evita depender del nombre específico del usuario de Windows.
  {: .lab-note .important .compact}

  ```powershell
  $LabRoot = Join-Path ([Environment]::GetFolderPath('Desktop')) 'finops-course'

  New-Item -ItemType Directory -Force -Path `
    "$LabRoot\data\raw", `
    "$LabRoot\data\processed", `
    "$LabRoot\outputs\lab9", `
    "$LabRoot\scripts" | Out-Null

  Set-Location $LabRoot
  ```

  > **Salida esperada:** Existen `data\raw`, `data\processed`, `outputs\lab9` y `scripts`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define las cuatro URL y descarga los datasets.

  > **Importante:** Sustituye únicamente los valores `PEGAR_AQUI_URL_DE_DESCARGA_*` por las URL definitivas proporcionadas por el instructor.
  {: .lab-note .important .compact}

  ```powershell
  $KubernetesUrl = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_kubernetes_costs_eur.csv"
  $AiUrl         = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_ai_workloads_eur.csv"
  $MulticloudUrl = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_multicloud_services.csv"
  $FxUrl         = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_fx_rates_eur.json"

  Invoke-WebRequest `
    -Uri $KubernetesUrl `
    -OutFile "$LabRoot\data\raw\novatech_kubernetes_costs_eur.csv"

  Invoke-WebRequest `
    -Uri $AiUrl `
    -OutFile "$LabRoot\data\raw\novatech_ai_workloads_eur.csv"

  Invoke-WebRequest `
    -Uri $MulticloudUrl `
    -OutFile "$LabRoot\data\raw\novatech_multicloud_services.csv"

  Invoke-WebRequest `
    -Uri $FxUrl `
    -OutFile "$LabRoot\data\raw\novatech_fx_rates_eur.json"
  ```

  > **Salida esperada:** Los cuatro archivos quedan disponibles en `data\raw`.
  {: .lab-note .output .compact}

### Tarea 1.2. Analizar costos y shared allocation

- {% include step_label.html %} Crea `scripts\modern_cost_drivers.py` y carga los cuatro archivos.

  > **Nota:** El script concentrará todo el análisis para mantener la práctica dentro de 15 minutos.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\modern_cost_drivers.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  from pathlib import Path
  import json
  import pandas as pd
  import numpy as np

  ROOT = Path.home() / "Desktop" / "finops-course"
  RAW = ROOT / "data" / "raw"
  PROCESSED = ROOT / "data" / "processed"

  k8s = pd.read_csv(
      RAW / "novatech_kubernetes_costs_eur.csv"
  )

  ai = pd.read_csv(
      RAW / "novatech_ai_workloads_eur.csv"
  )

  multicloud = pd.read_csv(
      RAW / "novatech_multicloud_services.csv"
  )

  with open(
      RAW / "novatech_fx_rates_eur.json",
      "r",
      encoding="utf-8"
  ) as f:
      fx = json.load(f)

  print(
      f"Kubernetes: {len(k8s):,} | "
      f"AI: {len(ai):,} | "
      f"Multicloud: {len(multicloud):,}"
  )
  ```

  > **Salida esperada:** Se muestran aproximadamente 1,694 registros Kubernetes, 465 AI y 930 Multicloud.
  {: .lab-note .output .compact}

- {% include step_label.html %} Identifica los principales drivers Kubernetes y compara tres métodos de shared allocation.

  > **Importante:** `platform-system` y `monitoring` representan el pool de costos compartidos que NovaTech distribuirá entre namespaces de aplicación.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\modern_cost_drivers.py`, agrega este bloque **debajo del bloque que imprime** `Kubernetes`, `AI` y `Multicloud`.
  {: .lab-note .important .compact}

  ```python
  infra_namespaces = [
      "platform-system",
      "monitoring"
  ]

  app_k8s = k8s[
      ~k8s["Namespace"].isin(infra_namespaces)
  ].copy()

  shared_pool_eur = k8s.loc[
      k8s["Namespace"].isin(infra_namespaces),
      "SharedCostEUR"
  ].sum()

  namespace_summary = (
      app_k8s.groupby("Namespace", as_index=False)
      .agg(
          DirectCostEUR=("DirectCostEUR", "sum"),
          CpuCoreHours=("CpuCoreHours", "sum")
      )
  )

  namespace_summary["EqualSharedEUR"] = (
      shared_pool_eur / len(namespace_summary)
  )

  namespace_summary["CpuSharedEUR"] = (
      namespace_summary["CpuCoreHours"]
      / namespace_summary["CpuCoreHours"].sum()
      * shared_pool_eur
  )

  namespace_summary["DirectCostSharedEUR"] = (
      namespace_summary["DirectCostEUR"]
      / namespace_summary["DirectCostEUR"].sum()
      * shared_pool_eur
  )

  namespace_summary["TotalWithCpuAllocationEUR"] = (
      namespace_summary["DirectCostEUR"]
      + namespace_summary["CpuSharedEUR"]
  )

  namespace_summary = namespace_summary.sort_values(
      "TotalWithCpuAllocationEUR",
      ascending=False
  )

  print(f"Shared pool: €{shared_pool_eur:,.2f}")
  print(namespace_summary.round(2).to_string(index=False))
  ```

  > **Salida esperada:** Se muestran los namespaces ordenados por costo y tres métodos de asignación del mismo shared pool.
  {: .lab-note .output .compact}

{% assign results = site.data.task-results[page.slug].results %}
{% capture r1 %}{{ results[0] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r1 %}

---

## 🌐 Tarea 2. Analizar Multicloud y AI — 6 min

En esta tarea normalizarás costos multicloud a EUR, identificarás variaciones de precio observadas y evaluarás la relación valor/costo de workloads de AI.

### Tarea 2.1. Normalizar Multicloud a EUR

- {% include step_label.html %} Convierte todos los costos multicloud a EUR y calcula costo por unidad.

  > **Nota:** Las tasas FX son fijas para que todos los participantes obtengan los mismos resultados.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\modern_cost_drivers.py`, agrega este bloque **debajo de** `print(namespace_summary.round(2).to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  rates = fx["rates_to_eur"]

  multicloud["NormalizedCostEUR"] = (
      multicloud["Cost"]
      * multicloud["Currency"].map(rates)
  )

  multicloud["CostPerUnitEUR"] = (
      multicloud["NormalizedCostEUR"]
      / multicloud["UnitsConsumed"]
  )

  mc_summary = (
      multicloud.groupby(
          ["ServiceCategory", "Provider"],
          as_index=False
      )
      .agg(
          TotalCostEUR=("NormalizedCostEUR", "sum"),
          AvgCostPerUnitEUR=("CostPerUnitEUR", "mean"),
          Units=("UnitsConsumed", "sum")
      )
  )

  print(
      mc_summary.round(4)
      .to_string(index=False)
  )
  ```

  > **Salida esperada:** Todos los proveedores quedan expresados en EUR con costo promedio por unidad.
  {: .lab-note .output .compact}

- {% include step_label.html %} Calcula la variación observada de precio por categoría.

  > **Importante:** Esta variación es una señal para análisis; no representa ahorro garantizado ni una recomendación automática de migración.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\modern_cost_drivers.py`, agrega este bloque **debajo del bloque que imprime** `mc_summary`.
  {: .lab-note .important .compact}

  ```python
  variance_rows = []

  for category, group in mc_summary.groupby(
      "ServiceCategory"
  ):
      low = group.loc[
          group["AvgCostPerUnitEUR"].idxmin()
      ]

      high = group.loc[
          group["AvgCostPerUnitEUR"].idxmax()
      ]

      variance_pct = (
          (
              high["AvgCostPerUnitEUR"]
              - low["AvgCostPerUnitEUR"]
          )
          / high["AvgCostPerUnitEUR"]
          * 100
      )

      variance_rows.append({
          "ServiceCategory": category,
          "LowestObservedProvider": low["Provider"],
          "HighestObservedProvider": high["Provider"],
          "LowestUnitCostEUR": round(
              low["AvgCostPerUnitEUR"], 4
          ),
          "HighestUnitCostEUR": round(
              high["AvgCostPerUnitEUR"], 4
          ),
          "IndicativeVariancePct": round(
              variance_pct, 1
          )
      })

  variance_df = pd.DataFrame(
      variance_rows
  ).sort_values(
      "IndicativeVariancePct",
      ascending=False
  )

  print(variance_df.to_string(index=False))
  ```

  > **Salida esperada:** Se identifican las categorías con mayor diferencia observada de costo unitario entre proveedores.
  {: .lab-note .output .compact}

### Tarea 2.2. Evaluar economics de AI

- {% include step_label.html %} Calcula costo por inferencia, Value/Cost Ratio y Net Value.

  > **Nota:** `AttributedBusinessValueEUR` es un supuesto didáctico para relacionar consumo tecnológico con valor de negocio.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\modern_cost_drivers.py`, agrega este bloque **debajo de** `print(variance_df.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  ai_summary = (
      ai.groupby(
          ["ModelName", "ModelType", "UseCase"],
          as_index=False
      )
      .agg(
          InferenceRequests=(
              "InferenceRequests", "sum"
          ),
          TokensProcessed=(
              "TokensProcessed", "sum"
          ),
          TotalCostEUR=(
              "TotalCostEUR", "sum"
          ),
          AttributedBusinessValueEUR=(
              "AttributedBusinessValueEUR", "sum"
          )
      )
  )

  ai_summary["CostPerInferenceEUR"] = (
      ai_summary["TotalCostEUR"]
      / ai_summary["InferenceRequests"]
  )

  ai_summary["CostPer1KTokensEUR"] = np.where(
      ai_summary["TokensProcessed"] > 0,
      (
          ai_summary["TotalCostEUR"]
          / ai_summary["TokensProcessed"]
          * 1000
      ),
      np.nan
  )

  ai_summary["ValueCostRatio"] = (
      ai_summary["AttributedBusinessValueEUR"]
      / ai_summary["TotalCostEUR"]
  )

  ai_summary["NetValueEUR"] = (
      ai_summary["AttributedBusinessValueEUR"]
      - ai_summary["TotalCostEUR"]
  )

  print(
      ai_summary[
          [
              "ModelName",
              "UseCase",
              "TotalCostEUR",
              "CostPerInferenceEUR",
              "CostPer1KTokensEUR",
              "ValueCostRatio",
              "NetValueEUR"
          ]
      ]
      .round(4)
      .sort_values("ValueCostRatio")
      .to_string(index=False)
  )
  ```

  > **Salida esperada:** Se observan workloads con diferentes relaciones entre costo y valor atribuido.
  {: .lab-note .output .compact}

- {% include step_label.html %} Selecciona los dos drivers AI con menor Value/Cost Ratio.

  > **Nota:** Un ratio bajo indica que el workload merece revisión de volumen, modelo, arquitectura o caso de uso.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\modern_cost_drivers.py`, agrega este bloque **debajo del bloque que imprime** `ai_summary`.
  {: .lab-note .important .compact}

  ```python
  ai_drivers = (
      ai_summary.sort_values(
          "ValueCostRatio"
      )
      .head(2)
      .copy()
  )

  print(
      ai_drivers[
          [
              "ModelName",
              "UseCase",
              "ValueCostRatio",
              "TotalCostEUR"
          ]
      ]
      .round(3)
      .to_string(index=False)
  )
  ```

  > **Salida esperada:** Se muestran los dos workloads de AI con menor relación valor/costo.
  {: .lab-note .output .compact}

{% capture r2 %}{{ results[1] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r2 %}

---

## 📋 Tarea 3. Enriquecer el backlog de optimización — 4 min

En esta tarea convertirás los principales drivers detectados en nuevas oportunidades compatibles con el backlog generado en la Práctica 6.

### Tarea 3.1. Crear seis oportunidades

- {% include step_label.html %} Selecciona dos drivers Kubernetes, dos Multicloud y dos AI.

  > **Importante:** Las nuevas oportunidades se crean para investigar y priorizar acciones; no representan cambios automáticos en infraestructura.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\modern_cost_drivers.py`, agrega este bloque **debajo del bloque que imprime** `ai_drivers`.
  {: .lab-note .important .compact}

  ```python
  k8s_drivers = namespace_summary.head(2).copy()

  mc_drivers = variance_df.head(2).copy()

  opportunities = []

  for _, row in k8s_drivers.iterrows():
      opportunities.append({
          "OptimizationCategory": "Kubernetes",
          "ActionType": "Efficiency",
          "Driver": row["Namespace"],
          "CurrentCostEUR": round(
              row["TotalWithCpuAllocationEUR"], 2
          ),
          "EstimatedMonthlySavingsEUR": round(
              row["TotalWithCpuAllocationEUR"] * 0.15,
              2
          ),
          "EffortDays": 3,
          "RiskLevel": "Medium",
          "RecommendedAction": (
              "Review rightsizing and shared-cost drivers"
          )
      })

  for _, row in mc_drivers.iterrows():
      opportunities.append({
          "OptimizationCategory": "Multicloud",
          "ActionType": "Architecture Review",
          "Driver": row["ServiceCategory"],
          "CurrentCostEUR": 0,
          "EstimatedMonthlySavingsEUR": 0,
          "EffortDays": 5,
          "RiskLevel": "Medium",
          "RecommendedAction": (
              "Investigate observed unit-cost variance "
              "and validate service comparability"
          )
      })

  for _, row in ai_drivers.iterrows():
      opportunities.append({
          "OptimizationCategory": "AI",
          "ActionType": "Value Optimization",
          "Driver": row["ModelName"],
          "CurrentCostEUR": round(
              row["TotalCostEUR"], 2
          ),
          "EstimatedMonthlySavingsEUR": round(
              row["TotalCostEUR"] * 0.10,
              2
          ),
          "EffortDays": 3,
          "RiskLevel": "Medium",
          "RecommendedAction": (
              "Review model, workload volume and "
              "value-to-cost economics"
          )
      })

  new_opportunities = pd.DataFrame(
      opportunities
  )

  print(new_opportunities.to_string(index=False))
  ```

  > **Salida esperada:** Se crean exactamente seis oportunidades: dos Kubernetes, dos Multicloud y dos AI.
  {: .lab-note .output .compact}

- {% include step_label.html %} Calcula el `PriorityScore` con la misma lógica utilizada en la Práctica 6.

  > **Nota:** Las oportunidades Multicloud mantienen ahorro 0 porque la variación observada requiere validación antes de estimar un ahorro realizable.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\modern_cost_drivers.py`, agrega este bloque **debajo de** `print(new_opportunities.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  risk_factor = {
      "Low": 1.0,
      "Medium": 0.7,
      "High": 0.4
  }

  new_opportunities["RiskFactor"] = (
      new_opportunities["RiskLevel"]
      .map(risk_factor)
  )

  new_opportunities["PriorityScore"] = (
      new_opportunities[
          "EstimatedMonthlySavingsEUR"
      ]
      / new_opportunities["EffortDays"]
      * new_opportunities["RiskFactor"]
  ).round(2)

  new_opportunities["Status"] = "Pending Review"
  ```

  > **Salida esperada:** Las seis oportunidades quedan preparadas para integrarse al backlog.
  {: .lab-note .output .compact}

### Tarea 3.2. Exportar el análisis y backlog enriquecido

- {% include step_label.html %} Integra las nuevas oportunidades al backlog existente y exporta los resultados.

  > **Importante:** La práctica reutiliza `finops_optimization_backlog.xlsx` generado en la Práctica 6.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\modern_cost_drivers.py`, agrega este bloque **debajo de** `new_opportunities["Status"] = "Pending Review"`.
  {: .lab-note .important .compact}

  ```python
  backlog_path = (
      PROCESSED
      / "finops_optimization_backlog.xlsx"
  )

  existing_backlog = pd.read_excel(
      backlog_path,
      sheet_name="backlog"
  )

  start_id = (
      int(existing_backlog["BacklogId"].max())
      + 1
  )

  new_opportunities.insert(
      0,
      "BacklogId",
      range(
          start_id,
          start_id + len(new_opportunities)
      )
  )

  enriched_backlog = pd.concat(
      [
          existing_backlog,
          new_opportunities
      ],
      ignore_index=True,
      sort=False
  )

  analysis_xlsx = (
      PROCESSED
      / "finops_modern_cost_drivers.xlsx"
  )

  enriched_xlsx = (
      PROCESSED
      / "finops_optimization_backlog_enriched.xlsx"
  )

  with pd.ExcelWriter(
      analysis_xlsx,
      engine="openpyxl"
  ) as writer:
      namespace_summary.to_excel(
          writer,
          sheet_name="kubernetes",
          index=False
      )

      variance_df.to_excel(
          writer,
          sheet_name="multicloud",
          index=False
      )

      ai_summary.to_excel(
          writer,
          sheet_name="ai",
          index=False
      )

      new_opportunities.to_excel(
          writer,
          sheet_name="new_opportunities",
          index=False
      )

  with pd.ExcelWriter(
      enriched_xlsx,
      engine="openpyxl"
  ) as writer:
      enriched_backlog.to_excel(
          writer,
          sheet_name="backlog",
          index=False
      )

  print(f"Creado: {analysis_xlsx}")
  print(f"Creado: {enriched_xlsx}")
  ```

  > **Salida esperada:** Se crean `finops_modern_cost_drivers.xlsx` y `finops_optimization_backlog_enriched.xlsx`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el script y valida los artefactos finales.

  > **Importante:** El archivo de análisis debe contener cuatro hojas y el backlog enriquecido debe incluir seis nuevas oportunidades.
  {: .lab-note .important .compact}

  ```powershell
  python "$LabRoot\scripts\modern_cost_drivers.py"
  ```
  ```powershell
  Get-Item `
    "$LabRoot\data\processed\finops_modern_cost_drivers.xlsx", `
    "$LabRoot\data\processed\finops_optimization_backlog_enriched.xlsx" |
    Select-Object Name, Length
  ```
  ```powershell
  python -c "import pandas as pd; p=r'$LabRoot\data\processed\finops_modern_cost_drivers.xlsx'; print(pd.ExcelFile(p).sheet_names)"
  ```

  > **Salida esperada:** Los dos Excel existen y el archivo de análisis contiene `kubernetes`, `multicloud`, `ai` y `new_opportunities`.
  {: .lab-note .output .compact}

{% capture r3 %}{{ results[2] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r3 %}
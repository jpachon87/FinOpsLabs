---
layout: lab
title: "Demostración 4: Asignación de costos por producto, equipo y ambiente"
permalink: /lab4/lab4/
images_base: /labs/lab4/img
duration: "30 minutos"
objective:
  - Implementar una estrategia básica de Allocation para clasificar, distribuir y reportar costos tecnológicos por equipo, producto y ambiente usando metadata organizacional.
prerequisites:
  - Equipo con Windows 11 y Visual Studio Code instalado.
  - Python 3.12 o superior disponible desde la terminal integrada de Visual Studio Code.
  - Paquetes pandas y openpyxl instalados.
  - Conocimiento básico de FinOps Allocation, showback, metadata organizacional y costos compartidos.
introduction:
  - En esta demostración trabajarás con un dataset de cuatro meses de NovaTech SRL preparado para Allocation. Clasificarás costos directos, compartidos, parciales y no asignados, definirás una taxonomía mínima de metadata, distribuirás shared costs y generarás un reporte de showback con métricas de calidad de asignación.
slug: lab4
lab_number: 4
final_result: >
  Al finalizar tendrás en el Escritorio un reporte Excel con resumen de asignación, costos por equipo, shared costs distribuidos, registros no asignados y métricas de calidad de Allocation para NovaTech SRL.
notes:
  - Todos los archivos se almacenan dentro de la carpeta finops-course del Escritorio de Windows.
  - El dataset usa EUR como BillingCurrency y EffectiveCost como métrica principal para Allocation.
  - La taxonomía de metadata utilizada en esta práctica pertenece a NovaTech y no representa un conjunto obligatorio definido por FinOps.
  - El prorrateo de shared costs mediante consumo de compute es una regla organizacional de NovaTech para esta demostración.
  - El porcentaje de cobertura y los objetivos de asignación son metas internas de NovaTech, no umbrales universales del FinOps Framework.
references:
  - text: "FinOps Allocation"
    url: https://www.finops.org/framework/capabilities/allocation/
  - text: "FinOps Invoicing & Chargeback"
    url: https://www.finops.org/framework/capabilities/invoicing-chargeback/
  - text: "FOCUS Specification"
    url: https://focus.finops.org/
  - text: "pandas GroupBy"
    url: https://pandas.pydata.org/docs/user_guide/groupby.html
prev: /lab3/lab3/
next: /lab5/lab5/
---

---

<!-- Aquí comienzan las instrucciones paso a paso de la práctica -->

## 🧭 Tarea 1. Preparar la estrategia de Allocation — 7 min

En esta tarea descargarás el dataset de Allocation, validarás su estructura y definirás la taxonomía de metadata que NovaTech utilizará para asignar costos por equipo, producto, ambiente y centro de costo.

### Tarea 1.1. Descargar y validar el dataset

Crearás la estructura de trabajo, descargarás el archivo de Allocation y comprobarás que contiene metadata suficiente para realizar la clasificación.

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
    "$LabRoot\outputs\lab4", `
    "$LabRoot\scripts" | Out-Null
  Set-Location $LabRoot
  ```

  > **Salida esperada:** Existen `data\raw`, `data\processed`, `outputs\lab4` y `scripts`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la URL del dataset y descárgalo en `data\raw`.

  ```powershell
  $DatasetUrl = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_allocation_4months_eur.csv"

  Invoke-WebRequest `
    -Uri $DatasetUrl `
    -OutFile "$LabRoot\data\raw\novatech_allocation_4months_eur.csv"
  ```

  > **Salida esperada:** Se crea `data\raw\novatech_allocation_4months_eur.csv`.
  {: .lab-note .output .compact}

### Tarea 1.2. Definir taxonomía y reglas de metadata

Validarás los campos principales del dataset y documentarás las cinco dimensiones de metadata utilizadas por NovaTech para Allocation.

- {% include step_label.html %} Verifica registros, período, moneda y columnas de metadata.

  > **Nota:** El dataset contiene registros completos, parciales, sin tags y shared services para ejercitar Allocation.
  {: .lab-note .info .compact}

  ```powershell
  python -c "import pandas as pd; p=r'$LabRoot\data\raw\novatech_allocation_4months_eur.csv'; df=pd.read_csv(p); print('Registros:',len(df)); print('Periodo:',df['ChargePeriodStart'].min(),'a',df['ChargePeriodStart'].max()); print('Monedas:',df['BillingCurrency'].unique().tolist()); print('Tags:',[c for c in df.columns if c.startswith('Tag')])"
  ```

  > **Salida esperada:** Se muestran aproximadamente 5,000 registros, período de octubre 2023 a enero 2024, EUR y cinco columnas `Tag...`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea `scripts\allocation_analysis.py` y define la taxonomía NovaTech.

  > **Importante:** Estas cinco dimensiones son una decisión organizacional de NovaTech y no un requisito universal de FinOps.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  from pathlib import Path
  import pandas as pd

  ROOT = Path.home() / "Desktop" / "finops-course"
  RAW = ROOT / "data" / "raw" / "novatech_allocation_4months_eur.csv"
  PROCESSED = ROOT / "data" / "processed"
  OUTPUTS = ROOT / "outputs" / "lab4"

  PROCESSED.mkdir(parents=True, exist_ok=True)
  OUTPUTS.mkdir(parents=True, exist_ok=True)

  df = pd.read_csv(RAW, parse_dates=["ChargePeriodStart"])

  taxonomy = pd.DataFrame({
      "TagKey": [
          "TagEnvironment",
          "TagTeam",
          "TagProduct",
          "TagCostCenter",
          "TagOwner"
      ],
      "Required": [True, True, True, True, False],
      "Purpose": [
          "Separar prod, staging y dev",
          "Asignar responsabilidad operativa",
          "Relacionar gasto con producto",
          "Relacionar gasto con estructura financiera",
          "Identificar contacto responsable"
      ]
  })

  print(taxonomy.to_string(index=False))
  ```

  > **Salida esperada:** El script define una taxonomía de cinco dimensiones, cuatro obligatorias y una opcional.
  {: .lab-note .output .compact}

{% assign results = site.data.task-results[page.slug].results %}
{% capture r1 %}{{ results[0] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r1 %}

---

## 🧮 Tarea 2. Clasificar y distribuir los costos — 13 min

En esta tarea clasificarás los registros según su asignabilidad, aplicarás inferencia básica por cuenta y distribuirás shared costs mediante una regla proporcional definida por NovaTech.

### Tarea 2.1. Clasificar direct, shared, partial y unallocated

Crearás reglas programáticas para separar los registros según la calidad de su metadata y el tipo de servicio.

- {% include step_label.html %} Agrega la función de clasificación de Allocation al script.

  > **Nota:** `partial` identifica registros con metadata incompleta; `unallocated` representa registros sin información suficiente para atribución directa.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, agrega este bloque **debajo de** `print(taxonomy.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  SHARED_SERVICES = {
      "AWS Transit Gateway",
      "AWS WAF",
      "AWS CloudTrail",
      "Azure Firewall",
      "Azure DNS",
      "Cloud DNS",
      "Cloud Armor"
  }

  def classify_allocation(row):
      if row["ServiceName"] in SHARED_SERVICES:
          return "shared"

      required = [
          row["TagEnvironment"],
          row["TagTeam"],
          row["TagProduct"],
          row["TagCostCenter"]
      ]

      has_any = any(pd.notna(v) and str(v).strip() for v in required)
      has_all = all(pd.notna(v) and str(v).strip() for v in required)

      if has_all:
          return "direct"
      if has_any:
          return "partial"
      return "unallocated"

  df["AllocationType"] = df.apply(classify_allocation, axis=1)
  ```

  > **Salida esperada:** Cada registro queda clasificado como `direct`, `shared`, `partial` o `unallocated`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega inferencia por `AccountId` para completar parte de la metadata faltante.

  > **Importante:** La inferencia se aplica solo cuando existe una relación organizacional conocida entre cuenta, equipo, producto y ambiente.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, agrega este bloque **debajo de** `df["AllocationType"] = df.apply(classify_allocation, axis=1)`.
  {: .lab-note .important .compact}

  ```python
  account_map = {
      "acc-paycore-prod": ("Backend", "PayCore", "prod", "CC-001"),
      "acc-paycore-dev": ("Backend", "PayCore", "dev", "CC-001"),
      "acc-paycore-staging": ("Backend", "PayCore", "staging", "CC-001"),
      "acc-devportal-prod": ("Frontend", "DevPortal", "prod", "CC-003"),
      "acc-devportal-dev": ("Frontend", "DevPortal", "dev", "CC-003"),
      "acc-devportal-staging": ("Frontend", "DevPortal", "staging", "CC-003"),
      "acc-analytics-prod": ("Data", "AnalyticsHub", "prod", "CC-002"),
      "acc-analytics-dev": ("Data", "AnalyticsHub", "dev", "CC-002"),
      "acc-analytics-staging": ("Data", "AnalyticsHub", "staging", "CC-002")
  }

  for account_id, values in account_map.items():
      mask = (df["AccountId"] == account_id) & (df["AllocationType"] == "partial")
      team, product, environment, cost_center = values

      df.loc[mask & df["TagTeam"].isna(), "TagTeam"] = team
      df.loc[mask & df["TagProduct"].isna(), "TagProduct"] = product
      df.loc[mask & df["TagEnvironment"].isna(), "TagEnvironment"] = environment
      df.loc[mask & df["TagCostCenter"].isna(), "TagCostCenter"] = cost_center

  df["AllocationType"] = df.apply(classify_allocation, axis=1)
  ```

  > **Salida esperada:** Parte de los registros `partial` se reclasifican como `direct`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Genera un resumen de registros y costo por tipo de Allocation.

  > **Nota:** La métrica se calcula con `EffectiveCost`, no con cantidad de registros únicamente.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, agrega este bloque **debajo de la segunda aparición de** `df["AllocationType"] = df.apply(classify_allocation, axis=1)`, después de aplicar la inferencia por `AccountId`.
  {: .lab-note .important .compact}

  ```python
  allocation_summary = (
      df.groupby("AllocationType", as_index=False)
        .agg(
            Records=("RecordId", "count"),
            EffectiveCostEUR=("EffectiveCost", "sum")
        )
  )

  allocation_summary["PctCost"] = (
      allocation_summary["EffectiveCostEUR"]
      / allocation_summary["EffectiveCostEUR"].sum()
      * 100
  ).round(1)

  print(allocation_summary.to_string(index=False))
  ```

  > **Salida esperada:** Se muestran los cuatro tipos de Allocation con registros, costo total y porcentaje del gasto.
  {: .lab-note .output .compact}

### Tarea 2.2. Distribuir shared costs con prorrateo proporcional

Aplicarás una regla de NovaTech que reparte shared costs según el EffectiveCost de servicios de compute directamente atribuibles a cada equipo.

- {% include step_label.html %} Define la base de prorrateo usando servicios de compute directos.

  > **Importante:** Esta regla de reparto pertenece a NovaTech; otras organizaciones pueden utilizar modelos diferentes.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, agrega este bloque **debajo de** `print(allocation_summary.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  COMPUTE_SERVICES = {
      "Amazon EC2",
      "AWS Lambda",
      "Azure Virtual Machines",
      "Azure Functions",
      "Compute Engine"
  }

  compute_by_team = (
      df[
          (df["AllocationType"] == "direct")
          & (df["ServiceName"].isin(COMPUTE_SERVICES))
      ]
      .groupby("TagTeam", as_index=False)
      .agg(ComputeCostEUR=("EffectiveCost", "sum"))
  )

  compute_by_team["ProrationPct"] = (
      compute_by_team["ComputeCostEUR"]
      / compute_by_team["ComputeCostEUR"].sum()
  )

  print(compute_by_team.to_string(index=False))
  ```

  > **Salida esperada:** Cada equipo obtiene un porcentaje de prorrateo cuya suma es 100%.
  {: .lab-note .output .compact}

- {% include step_label.html %} Distribuye el total de shared costs según la base calculada.

  > **Nota:** El total distribuido debe reconciliar exactamente con el shared cost original, salvo diferencias menores de redondeo.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, agrega este bloque **debajo de** `print(compute_by_team.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  shared_total = df.loc[
      df["AllocationType"] == "shared",
      "EffectiveCost"
  ].sum()

  shared_allocation = compute_by_team[["TagTeam", "ProrationPct"]].copy()
  shared_allocation["SharedCostAllocatedEUR"] = (
      shared_allocation["ProrationPct"] * shared_total
  )

  print(shared_allocation.to_string(index=False))
  print(
      "Shared original:",
      round(shared_total, 2),
      "Distribuido:",
      round(shared_allocation["SharedCostAllocatedEUR"].sum(), 2)
  )
  ```

  > **Salida esperada:** La suma del shared cost distribuido coincide con el total compartido original.
  {: .lab-note .output .compact}

- {% include step_label.html %} Construye la tabla consolidada de Allocation por equipo.

  > **Nota:** Los costos `partial` y `unallocated` permanecen separados y no se fuerzan artificialmente a un equipo.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, agrega este bloque **debajo del bloque que imprime** `Shared original` y `Distribuido`.
  {: .lab-note .important .compact}

  ```python
  direct_by_team = (
      df[df["AllocationType"] == "direct"]
      .groupby("TagTeam", as_index=False)
      .agg(DirectCostEUR=("EffectiveCost", "sum"))
  )

  allocation_by_team = direct_by_team.merge(
      shared_allocation[["TagTeam", "SharedCostAllocatedEUR"]],
      on="TagTeam",
      how="left"
  ).fillna(0)

  allocation_by_team["AllocatedTotalEUR"] = (
      allocation_by_team["DirectCostEUR"]
      + allocation_by_team["SharedCostAllocatedEUR"]
  )

  print(allocation_by_team.to_string(index=False))
  ```

  > **Salida esperada:** Se obtiene una tabla por equipo con costo directo, shared cost distribuido y total asignado.
  {: .lab-note .output .compact}

{% capture r2 %}{{ results[1] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r2 %}

---

## 📊 Tarea 3. Generar showback y métricas de calidad — 10 min

En esta tarea crearás un reporte de showback y calcularás indicadores de cobertura, shared cost y unallocated cost para evaluar la calidad del modelo de Allocation de NovaTech.

### Tarea 3.1. Construir el reporte de showback

Generarás una vista informativa por equipo y ambiente, manteniendo separados los costos que todavía no cuentan con atribución confiable.

- {% include step_label.html %} Crea el desglose de costos directos por equipo y ambiente.

  > **Nota:** Showback proporciona visibilidad del consumo sin implicar necesariamente un cargo contable interno.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, agrega este bloque **debajo de** `print(allocation_by_team.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  direct_detail = (
      df[df["AllocationType"] == "direct"]
      .pivot_table(
          index="TagTeam",
          columns="TagEnvironment",
          values="EffectiveCost",
          aggfunc="sum",
          fill_value=0
      )
      .reset_index()
  )

  print(direct_detail.to_string(index=False))
  ```

  > **Salida esperada:** Se muestra el EffectiveCost por equipo para prod, staging y dev.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea la vista de showback por equipo.

  > **Importante:** Los costos `partial` y `unallocated` no se redistribuyen de forma arbitraria.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, agrega este bloque **debajo de** `print(direct_detail.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  showback = allocation_by_team.copy()
  showback["Model"] = "Showback"
  showback["ChargeApplied"] = False

  unallocated_cost = df[
      df["AllocationType"].isin(["partial", "unallocated"])
  ]["EffectiveCost"].sum()

  print(showback.to_string(index=False))
  print(f"\nUnallocated / partial separado: €{unallocated_cost:,.2f}")
  ```

  > **Salida esperada:** El showback muestra costos por equipo y mantiene separado el costo no atribuido.
  {: .lab-note .output .compact}

- {% include step_label.html %} Calcula las métricas de calidad de Allocation.

  > **Nota:** Los objetivos mostrados son metas internas de NovaTech y no umbrales universales de FinOps.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, agrega este bloque **debajo de** `print(f"\nUnallocated / partial separado: €{unallocated_cost:,.2f}")`.
  {: .lab-note .important .compact}

  ```python
  total_cost = df["EffectiveCost"].sum()
  direct_cost = df.loc[
      df["AllocationType"] == "direct",
      "EffectiveCost"
  ].sum()

  shared_cost = df.loc[
      df["AllocationType"] == "shared",
      "EffectiveCost"
  ].sum()

  residual_cost = df.loc[
      df["AllocationType"].isin(["partial", "unallocated"]),
      "EffectiveCost"
  ].sum()

  metrics = pd.DataFrame({
      "Metric": [
          "DirectAllocationPct",
          "SharedCostPct",
          "ResidualUnallocatedPct",
          "DirectCostEUR",
          "SharedCostEUR",
          "ResidualCostEUR"
      ],
      "Value": [
          round(direct_cost / total_cost * 100, 1),
          round(shared_cost / total_cost * 100, 1),
          round(residual_cost / total_cost * 100, 1),
          round(direct_cost, 2),
          round(shared_cost, 2),
          round(residual_cost, 2)
      ],
      "NovaTechTarget": [
          ">= 80%",
          "100% distributed",
          "<= 5%",
          "Maximize",
          "Reconcile",
          "Minimize"
      ]
  })

  print(metrics.to_string(index=False))
  ```

  > **Salida esperada:** Se muestran porcentajes y valores de costo directo, shared y residual no asignado.
  {: .lab-note .output .compact}

### Tarea 3.2. Exportar y validar el reporte final

Exportarás los principales resultados a Excel y comprobarás que el reporte contiene todas las vistas necesarias.

- {% include step_label.html %} Exporta `novatech_allocation_report.xlsx` con cuatro hojas.

  > **Nota:** El reporte consolida Allocation, shared cost, showback y registros pendientes de atribución.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\allocation_analysis.py`, agrega este bloque **debajo de** `print(metrics.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  report_xlsx = PROCESSED / "novatech_allocation_report.xlsx"

  residual = (
      df.loc[
          df["AllocationType"].isin(["partial", "unallocated"]),
          [
              "RecordId",
              "ChargePeriodStart",
              "ServiceProviderName",
              "AccountId",
              "ServiceName",
              "EffectiveCost",
              "BillingCurrency",
              "AllocationType"
          ]
      ]
      .sort_values("EffectiveCost", ascending=False)
  )

  with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
      metrics.to_excel(writer, sheet_name="summary", index=False)
      showback.to_excel(writer, sheet_name="allocation_by_team", index=False)
      shared_allocation.to_excel(writer, sheet_name="shared_cost", index=False)
      residual.to_excel(writer, sheet_name="unallocated", index=False)

  print(f"Reporte creado: {report_xlsx}")
  ```

  > **Salida esperada:** Se crea `data\processed\novatech_allocation_report.xlsx` con cuatro hojas.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el script completo y valida el archivo generado.

  > **Importante:** El reporte debe existir, contener datos y preservar por separado los costos no asignados.
  {: .lab-note .important .compact}

  ```powershell
  python "$LabRoot\scripts\allocation_analysis.py"
  ```
  ```powershell
  Get-Item "$LabRoot\data\processed\novatech_allocation_report.xlsx" |
    Select-Object Name, Length
  ```
  ```powershell
  python -c "import pandas as pd; p=r'$LabRoot\data\processed\novatech_allocation_report.xlsx'; x=pd.ExcelFile(p); print(x.sheet_names)"
  ```

  > **Salida esperada:** El archivo tiene tamaño mayor a 0 bytes y contiene `summary`, `allocation_by_team`, `shared_cost` y `unallocated`.
  {: .lab-note .output .compact}

{% capture r3 %}{{ results[2] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r3 %}
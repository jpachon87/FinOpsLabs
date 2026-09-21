---
layout: lab
title: "Demostración 2: Lectura de un reporte de consumo cloud"
permalink: /lab2/lab2/
images_base: /labs/lab2/img
duration: "25 minutos"
objective:
  - Interpretar un reporte de facturación cloud multicloud, diferenciar métricas de costo y generar un resumen financiero básico aplicando conceptos FinOps.
prerequisites:
  - Equipo con Windows 11 y Visual Studio Code instalado.
  - Python 3.12 o superior disponible desde la terminal integrada de Visual Studio Code.
  - Paquetes pandas, numpy y openpyxl instalados.
  - Conocimiento básico de pandas, tipos de costo cloud y modelos IaaS/PaaS/SaaS.
introduction:
  - En esta demostración trabajarás con un dataset sintético de facturación multicloud de NovaTech SRL correspondiente a enero de 2024. Cargarás y explorarás el reporte, identificarás campos de costo, revisarás la calidad de los datos, calcularás métricas por proveedor y servicio y exportarás un resumen consolidado en Excel.
slug: lab2
lab_number: 2
final_result: >
  Al finalizar tendrás un dataset de facturación multicloud con 500 registros, un script Python de análisis y un archivo Excel con resúmenes por proveedor, servicio, modelo de consumo y cobertura de compromisos.
notes:
  - Todos los archivos se almacenan dentro de la carpeta finops-course del Escritorio de Windows.
  - El dataset es sintético y representa a la organización ficticia NovaTech SRL.
  - Los campos de costo se utilizan con fines didácticos para practicar conceptos FinOps y normalización.
  - El dataset se inspira en terminología FOCUS, pero no representa una implementación completa de la especificación.
references:
  - text: "FOCUS Specification"
    url: https://focus.finops.org/
  - text: "FinOps Data Ingestion"
    url: https://www.finops.org/framework/capabilities/data-ingestion/
  - text: "pandas GroupBy"
    url: https://pandas.pydata.org/docs/user_guide/groupby.html
prev: /lab1/lab1/
next: /lab3/lab3/
---

---

<!-- Aquí comienzan las instrucciones paso a paso de la práctica -->

## 🔎 Tarea 1. Preparar el dataset y explorar el reporte — 7 min

En esta tarea crearás la estructura local de trabajo, generarás el dataset sintético de NovaTech SRL y cargarás el reporte en un Jupyter script para revisar sus campos y estructura.

### Tarea 1.1. Preparar el entorno y generar el dataset

- {% include step_label.html %} Abre Visual Studio Code y una terminal integrada de PowerShell.

  > **Nota:** Usa **Terminal → New Terminal** y confirma que el perfil activo sea PowerShell.
  {: .lab-note .info .compact}

  > **Salida esperada:** VS Code muestra una terminal PowerShell disponible para ejecutar comandos.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la carpeta del laboratorio y crea la estructura necesaria.

  > **Importante:** La variable `$LabRoot` evita depender del nombre específico del usuario de Windows.
  {: .lab-note .important .compact}

  ```powershell
  $LabRoot = Join-Path ([Environment]::GetFolderPath('Desktop')) 'finops-course'

  New-Item -ItemType Directory -Force -Path `
    "$LabRoot\data\raw", `
    "$LabRoot\data\processed", `
    "$LabRoot\scripts", `
    "$LabRoot\outputs", `
    "$LabRoot\scripts" | Out-Null

  Set-Location $LabRoot
  ```

  > **Salida esperada:** Existen las carpetas `data\raw`, `data\processed`, `scripts`, `outputs` y `scripts`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea el archivo `scripts\generate_billing_data.py` desde el explorador de Visual Studio Code.

  > **Salida esperada:** Existe el archivo `scripts\generate_billing_data.py`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega el generador del dataset sintético.

  > **Ubicación del código:** En `scripts\generate_billing_data.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  import pandas as pd
  import numpy as np
  from datetime import datetime, timedelta
  from pathlib import Path

  np.random.seed(42)

  ROOT = Path.home() / "Desktop" / "finops-course"
  OUTPUT = ROOT / "data" / "raw" / "novatech_billing_raw.csv"
  OUTPUT.parent.mkdir(parents=True, exist_ok=True)

  providers = ["AWS", "Azure", "GCP"]
  provider_weights = [0.50, 0.30, 0.20]

  services_by_provider = {
      "AWS": {
          "IaaS": ["Amazon EC2", "Amazon EBS"],
          "PaaS": ["Amazon RDS", "AWS Lambda", "Amazon S3"],
          "SaaS": ["Amazon WorkSpaces"],
      },
      "Azure": {
          "IaaS": ["Azure Virtual Machines", "Azure Disk Storage"],
          "PaaS": ["Azure SQL Database", "Azure Functions", "Azure Blob Storage"],
          "SaaS": ["Microsoft 365"],
      },
      "GCP": {
          "IaaS": ["Compute Engine", "Persistent Disk"],
          "PaaS": ["Cloud SQL", "Cloud Functions", "Cloud Storage"],
          "SaaS": ["Google Workspace"],
      },
  }

  products = ["PayCore", "AnalyticsHub", "DevPortal"]
  teams = ["Backend", "Frontend", "Data", "Platform"]
  environments = ["prod", "staging", "dev"]
  env_weights = [0.60, 0.25, 0.15]

  regions = {
      "AWS": ["us-east-1", "eu-west-1", "sa-east-1"],
      "Azure": ["eastus", "westeurope", "brazilsouth"],
      "GCP": ["us-central1", "europe-west1", "southamerica-east1"],
  }

  charge_types = ["Usage", "Usage", "Usage", "Usage", "Purchase", "Tax", "Credit"]
  usage_types = ["OnDemand", "OnDemand", "OnDemand", "Reserved", "Reserved", "SpotInstance"]

  records = []
  start_date = datetime(2024, 1, 1)

  for i in range(500):
      provider = np.random.choice(providers, p=provider_weights)
      model = np.random.choice(["IaaS", "PaaS", "SaaS"], p=[0.45, 0.40, 0.15])
      service = np.random.choice(services_by_provider[provider][model])
      billing_date = start_date + timedelta(days=np.random.randint(0, 31))
      product = np.random.choice(products)
      team = np.random.choice(teams)
      environment = np.random.choice(environments, p=env_weights)
      region = np.random.choice(regions[provider])
      charge_type = np.random.choice(charge_types)
      usage_type = np.random.choice(usage_types) if charge_type == "Usage" else "N/A"

      if model == "IaaS":
          base_cost = np.random.uniform(5.0, 250.0)
      elif model == "PaaS":
          base_cost = np.random.uniform(2.0, 180.0)
      else:
          base_cost = np.random.uniform(10.0, 50.0)

      if environment == "prod":
          base_cost *= 1.5
      elif environment == "staging":
          base_cost *= 0.7
      else:
          base_cost *= 0.4

      billed_cost = round(base_cost, 4)

      if usage_type == "Reserved":
          effective_cost = round(billed_cost * 0.65, 4)
      elif usage_type == "SpotInstance":
          effective_cost = round(billed_cost * 0.30, 4)
      else:
          effective_cost = billed_cost

      if charge_type == "Purchase":
          amortized_cost = round(billed_cost / 12, 4)
      elif usage_type == "Reserved":
          amortized_cost = round(effective_cost * 1.05, 4)
      else:
          amortized_cost = effective_cost

      if charge_type == "Credit":
          net_cost = round(-abs(billed_cost) * 0.1, 4)
          billed_cost = net_cost
          effective_cost = net_cost
          amortized_cost = net_cost
      elif charge_type == "Tax":
          net_cost = round(billed_cost * 0.21, 4)
          billed_cost = net_cost
          effective_cost = net_cost
          amortized_cost = net_cost
      else:
          net_cost = amortized_cost

      resource_id = f"{provider.lower()}-{service.lower().replace(' ', '-')[:10]}-{i:04d}"

      records.append({
          "InvoiceId": f"INV-2024-01-{provider[:3]}-{i:04d}",
          "BillingPeriodStart": "2024-01-01",
          "BillingPeriodEnd": "2024-01-31",
          "BillingDate": billing_date.strftime("%Y-%m-%d"),
          "Provider": provider,
          "ServiceName": service,
          "ServiceModel": model,
          "ResourceId": resource_id,
          "Region": region,
          "ChargeType": charge_type,
          "UsageType": usage_type,
          "UsageQuantity": round(np.random.uniform(1, 744), 2),
          "UsageUnit": "Hours" if model == "IaaS" else (
              "Requests" if "Lambda" in service or "Functions" in service else "GB-Month"
          ),
          "BilledCost": billed_cost,
          "EffectiveCost": effective_cost,
          "AmortizedCost": amortized_cost,
          "NetCost": net_cost,
          "Currency": "USD",
          "TagProduct": product if np.random.random() > 0.05 else "",
          "TagTeam": team if np.random.random() > 0.05 else "",
          "TagEnvironment": environment if np.random.random() > 0.03 else "",
      })

  df = pd.DataFrame(records)
  df.to_csv(OUTPUT, index=False)

  print(f"Dataset generado: {len(df)} registros")
  print(f"Proveedores: {df['Provider'].value_counts().to_dict()}")
  print(f"Modelos: {df['ServiceModel'].value_counts().to_dict()}")
  print(f"Archivo guardado en: {OUTPUT}")
  ```

  > **Salida esperada:** El script queda listo para generar 500 registros de facturación multicloud.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el generador y valida el archivo creado.

  ```powershell
  python "$LabRoot\scripts\generate_billing_data.py"

  Get-Item "$LabRoot\data\raw\novatech_billing_raw.csv" |
    Select-Object Name, Length
  ```

  > **Salida esperada:** Se crea `novatech_billing_raw.csv` con 500 registros y tamaño mayor a 0 bytes.
  {: .lab-note .output .compact}

### Tarea 1.2. Crear el script de análisis y cargar el reporte

- {% include step_label.html %} Crea el archivo `scripts\billing_analysis.py` desde el explorador de Visual Studio Code.

  > **Salida esperada:** Existe el archivo `scripts\billing_analysis.py`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega las importaciones y rutas de trabajo.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  import pandas as pd
  import numpy as np
  from pathlib import Path

  pd.set_option("display.max_columns", None)
  pd.set_option("display.float_format", "{:.2f}".format)

  ROOT = Path.home() / "Desktop" / "finops-course"
  DATA_RAW = ROOT / "data" / "raw"
  DATA_PROCESSED = ROOT / "data" / "processed"
  OUTPUTS = ROOT / "outputs"

  DATA_PROCESSED.mkdir(parents=True, exist_ok=True)
  OUTPUTS.mkdir(parents=True, exist_ok=True)

  print(f"pandas: {pd.__version__}")
  print(f"numpy: {np.__version__}")
  ```

  > **Salida esperada:** El script queda preparado con las librerías y rutas de trabajo.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega la carga y exploración inicial del dataset.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **debajo de** `print(f"numpy: {np.__version__}")`.
  {: .lab-note .important .compact}

  ```python
  df = pd.read_csv(DATA_RAW / "novatech_billing_raw.csv")

  print(f"\nRegistros cargados: {df.shape[0]}")
  print(f"Columnas: {df.shape[1]}")
  print("\nColumnas disponibles:")

  for i, col in enumerate(df.columns, 1):
      print(f"{i:2d}. {col}")

  print("\nPrimeros 5 registros:")
  print(df.head().to_string(index=False))
  ```

  > **Salida esperada:** Se muestran 500 registros, 21 columnas y una vista previa del reporte.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el script desde PowerShell.

  ```powershell
  python "$LabRoot\scripts\billing_analysis.py"
  ```

  > **Salida esperada:** La terminal muestra las versiones de pandas y numpy, 500 registros, 21 columnas y los primeros registros del dataset.
  {: .lab-note .output .compact}

{% assign results = site.data.task-results[page.slug].results %}
{% capture r1 %}{{ results[0] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r1 %}

---

## 📊 Tarea 2. Interpretar, normalizar y analizar los costos — 10 min

En esta tarea revisarás los campos financieros del reporte, identificarás problemas básicos de calidad y calcularás métricas de costo por proveedor, servicio y modelo.

### Tarea 2.1. Interpretar tipos de costo y calidad de datos

- {% include step_label.html %} Analiza la distribución de tipos de cargo y uso.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **debajo de** `print(df.head().to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  print("DISTRIBUCIÓN POR TIPO DE CARGO")
  print("=" * 60)

  charge_dist = df["ChargeType"].value_counts()
  charge_pct = (charge_dist / len(df) * 100).round(1)

  for charge, count in charge_dist.items():
      print(f"{charge:<12} → {count:>4} registros ({charge_pct[charge]}%)")

  print("\nDISTRIBUCIÓN POR TIPO DE USO")
  print("=" * 60)

  usage_dist = df[df["ChargeType"] == "Usage"]["UsageType"].value_counts()
  usage_pct = (usage_dist / usage_dist.sum() * 100).round(1)

  for usage, count in usage_dist.items():
      print(f"{usage:<14} → {count:>4} registros ({usage_pct[usage]}%)")
  ```

  > **Salida esperada:** Se muestran las distribuciones de `Usage`, `Purchase`, `Tax`, `Credit` y los tipos de uso OnDemand, Reserved y SpotInstance.
  {: .lab-note .output .compact}

- {% include step_label.html %} Compara BilledCost, EffectiveCost, AmortizedCost y NetCost.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **debajo del bloque de distribución por tipo de uso**.
  {: .lab-note .important .compact}

  ```python
  usage_df = df[df["ChargeType"] == "Usage"].copy()

  cost_comparison = pd.DataFrame({
      "Métrica": ["BilledCost", "EffectiveCost", "AmortizedCost", "NetCost"],
      "Total_USD": [
          usage_df["BilledCost"].sum(),
          usage_df["EffectiveCost"].sum(),
          usage_df["AmortizedCost"].sum(),
          usage_df["NetCost"].sum(),
      ],
  })

  cost_comparison["Diferencia_vs_Billed_%"] = (
      (cost_comparison["Total_USD"] - cost_comparison["Total_USD"].iloc[0])
      / cost_comparison["Total_USD"].iloc[0]
      * 100
  ).round(2)

  cost_comparison
  ```

  > **Salida esperada:** `EffectiveCost` es menor que `BilledCost` cuando existen descuentos por reservas o Spot.
  {: .lab-note .output .compact}

- {% include step_label.html %} Normaliza tags, tipos numéricos y fechas.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **debajo de** `cost_comparison`.
  {: .lab-note .important .compact}

  ```python
  tag_cols = ["TagProduct", "TagTeam", "TagEnvironment"]
  cost_cols = [
      "BilledCost",
      "EffectiveCost",
      "AmortizedCost",
      "NetCost",
      "UsageQuantity",
  ]

  df_clean = df.copy()
  df_clean[tag_cols] = df_clean[tag_cols].replace("", np.nan)

  for col in cost_cols:
      df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce")

  df_clean["BillingDate"] = pd.to_datetime(df_clean["BillingDate"])
  df_clean["BillingPeriodStart"] = pd.to_datetime(df_clean["BillingPeriodStart"])
  df_clean["BillingPeriodEnd"] = pd.to_datetime(df_clean["BillingPeriodEnd"])
  df_clean["BillingDay"] = df_clean["BillingDate"].dt.day

  print(f"Registros finales: {len(df_clean)}")
  print(f"BillingDate: {df_clean['BillingDate'].dtype}")
  print(f"BilledCost: {df_clean['BilledCost'].dtype}")

  for col in tag_cols:
      empty_count = df_clean[col].isna().sum()
      coverage = ((len(df_clean) - empty_count) / len(df_clean) * 100)
      print(f"{col}: {coverage:.1f}% de cobertura")
  ```

  > **Salida esperada:** Permanecen 500 registros, `BillingDate` se reconoce como tipo datetime y los tags presentan aproximadamente 95–97% de cobertura.
  {: .lab-note .output .compact}

### Tarea 2.2. Calcular métricas por proveedor, servicio y modelo

- {% include step_label.html %} Calcula el resumen de costos por proveedor.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **debajo del bloque de normalización de `df_clean`**.
  {: .lab-note .important .compact}

  ```python
  provider_summary = df_clean.groupby("Provider").agg(
      Total_AmortizedCost=("AmortizedCost", "sum"),
      Total_BilledCost=("BilledCost", "sum"),
      Num_Registros=("InvoiceId", "count"),
      Costo_Promedio_Diario=("AmortizedCost", lambda x: x.sum() / 31),
  ).round(2)

  provider_summary["Participacion_%"] = (
      provider_summary["Total_AmortizedCost"]
      / provider_summary["Total_AmortizedCost"].sum()
      * 100
  ).round(1)

  provider_summary = provider_summary.sort_values(
      "Total_AmortizedCost",
      ascending=False
  )

  provider_summary
  ```

  > **Salida esperada:** Se obtiene una tabla comparativa de AWS, Azure y GCP con costo total, registros, promedio diario y participación.
  {: .lab-note .output .compact}

- {% include step_label.html %} Calcula los resúmenes por servicio y modelo.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **debajo de** `provider_summary`.
  {: .lab-note .important .compact}

  ```python
  service_summary = df_clean.groupby(
      ["ServiceName", "ServiceModel", "Provider"]
  ).agg(
      Total_AmortizedCost=("AmortizedCost", "sum"),
      Num_Registros=("InvoiceId", "count"),
      Costo_Promedio=("AmortizedCost", "mean"),
  ).round(2)

  service_summary["Participacion_%"] = (
      service_summary["Total_AmortizedCost"]
      / service_summary["Total_AmortizedCost"].sum()
      * 100
  ).round(1)

  service_summary = service_summary.sort_values(
      "Total_AmortizedCost",
      ascending=False
  )

  model_summary = df_clean.groupby("ServiceModel").agg(
      Total_AmortizedCost=("AmortizedCost", "sum"),
      Total_BilledCost=("BilledCost", "sum"),
      Num_Registros=("InvoiceId", "count"),
  ).round(2)

  model_summary["Ahorro_vs_Billed"] = (
      model_summary["Total_BilledCost"]
      - model_summary["Total_AmortizedCost"]
  ).round(2)

  model_summary["Participacion_%"] = (
      model_summary["Total_AmortizedCost"]
      / model_summary["Total_AmortizedCost"].sum()
      * 100
  ).round(1)

  print("TOP 10 SERVICIOS")
  print(service_summary.head(10).to_string())

  print("\nMODELOS DE SERVICIO")
  print(model_summary.to_string())
  ```

  > **Salida esperada:** Se muestran los diez servicios de mayor costo y un resumen separado de IaaS, PaaS y SaaS.
  {: .lab-note .output .compact}

- {% include step_label.html %} Analiza cobertura y ahorro de OnDemand, Reserved y SpotInstance.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **debajo del resumen por servicio y modelo**.
  {: .lab-note .important .compact}

  ```python
  usage_only = df_clean[df_clean["ChargeType"] == "Usage"].copy()

  coverage = usage_only.groupby("UsageType").agg(
      Total_BilledCost=("BilledCost", "sum"),
      Total_EffectiveCost=("EffectiveCost", "sum"),
      Num_Registros=("InvoiceId", "count"),
  ).round(2)

  coverage["Ahorro_USD"] = (
      coverage["Total_BilledCost"] - coverage["Total_EffectiveCost"]
  ).round(2)

  coverage["Descuento_%"] = (
      coverage["Ahorro_USD"] / coverage["Total_BilledCost"] * 100
  ).round(1)

  coverage
  ```

  > **Salida esperada:** OnDemand presenta 0% de descuento, Reserved aproximadamente 35% y SpotInstance aproximadamente 70%.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta nuevamente el script para validar el análisis financiero.

  ```powershell
  python "$LabRoot\scripts\billing_analysis.py"
  ```

  > **Salida esperada:** La terminal muestra las distribuciones, comparaciones de costo, resúmenes por proveedor, servicio, modelo y cobertura.
  {: .lab-note .output .compact}

{% capture r2 %}{{ results[1] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r2 %}

---

## 🚀 Tarea 3. Exportar y validar el reporte consolidado — 8 min

En esta tarea exportarás los resultados a Excel, comprobarás la integridad del archivo y generarás un resumen ejecutivo del análisis.

### Tarea 3.1. Generar el archivo Excel

- {% include step_label.html %} Exporta los resultados a `novatech_billing_summary_jan2024.xlsx`.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **debajo del análisis de cobertura**.
  {: .lab-note .important .compact}

  ```python
  output_file = OUTPUTS / "novatech_billing_summary_jan2024.xlsx"
  OUTPUTS.mkdir(parents=True, exist_ok=True)

  metadata = pd.DataFrame({
      "Campo": [
          "Organización",
          "Período",
          "Total Registros",
          "Proveedores",
          "Fecha Análisis",
          "Analista",
          "Costo Total Amortizado",
          "Moneda",
      ],
      "Valor": [
          "NovaTech SRL",
          "Enero 2024",
          str(len(df_clean)),
          "AWS, Azure, GCP",
          pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
          "Equipo FinOps",
          f"${df_clean['AmortizedCost'].sum():,.2f}",
          "USD",
      ],
  })

  with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
      provider_summary.to_excel(writer, sheet_name="Por_Proveedor")
      service_summary.reset_index().to_excel(
          writer,
          sheet_name="Por_Servicio",
          index=False
      )
      model_summary.to_excel(writer, sheet_name="Por_Modelo")
      coverage.to_excel(writer, sheet_name="Cobertura_Reservas")
      df_clean.to_excel(writer, sheet_name="Datos_Completos", index=False)
      metadata.to_excel(writer, sheet_name="Metadata", index=False)

  print(f"Archivo creado: {output_file}")
  ```

  > **Salida esperada:** Se crea `outputs\novatech_billing_summary_jan2024.xlsx` con seis hojas.
  {: .lab-note .output .compact}

- {% include step_label.html %} Valida la integridad del archivo generado.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **debajo de** `print(f"Archivo creado: {output_file}")`.
  {: .lab-note .important .compact}

  ```python
  verify = pd.ExcelFile(output_file)

  print(f"Hojas: {verify.sheet_names}")
  print(f"Total de hojas: {len(verify.sheet_names)}")

  df_verify = pd.read_excel(
      output_file,
      sheet_name="Datos_Completos"
  )

  print(f"Registros en Datos_Completos: {len(df_verify)}")
  ```

  > **Salida esperada:** El archivo contiene exactamente seis hojas y `Datos_Completos` conserva 500 registros.
  {: .lab-note .output .compact}

### Tarea 3.2. Generar el resumen ejecutivo y validación final

- {% include step_label.html %} Genera un resumen ejecutivo del reporte de facturación.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **debajo de la validación del archivo Excel**.
  {: .lab-note .important .compact}

  ```python
  total_amortized = df_clean["AmortizedCost"].sum()
  days_in_period = 31

  print("=" * 60)
  print("RESUMEN EJECUTIVO - REPORTE DE FACTURACIÓN")
  print("NovaTech SRL | Enero 2024")
  print("=" * 60)

  print(f"Costo Total Amortizado: ${total_amortized:,.2f}")
  print(f"Costo Promedio Diario:  ${total_amortized / days_in_period:,.2f}")
  print(f"Registros procesados:   {len(df_clean):,}")
  print(f"Proveedores activos:    {df_clean['Provider'].nunique()}")
  print(f"Servicios únicos:       {df_clean['ServiceName'].nunique()}")

  print("\nDistribución por modelo:")
  for model in ["IaaS", "PaaS", "SaaS"]:
      model_cost = df_clean.loc[
          df_clean["ServiceModel"] == model,
          "AmortizedCost"
      ].sum()
      model_pct = model_cost / total_amortized * 100
      print(f"{model}: ${model_cost:,.2f} ({model_pct:.1f}%)")
  ```

  > **Salida esperada:** Se muestran costo total amortizado, promedio diario, registros, proveedores, servicios y distribución por modelo.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta la validación final del laboratorio.

  > **Ubicación del código:** En `scripts\billing_analysis.py`, agrega este bloque **al final del archivo**.
  {: .lab-note .important .compact}

  ```python
  checks = [
      ("Dataset con 500 registros", len(df_clean) == 500),
      ("Columnas esperadas (22)", len(df_clean.columns) == 22),
      ("BillingDate es datetime", pd.api.types.is_datetime64_any_dtype(df_clean["BillingDate"])),
      ("BilledCost es numérico", pd.api.types.is_numeric_dtype(df_clean["BilledCost"])),
      ("Excel generado existe", output_file.exists()),
      ("Excel tiene 6 hojas", len(pd.ExcelFile(output_file).sheet_names) == 6),
      ("3 proveedores", df_clean["Provider"].nunique() == 3),
      ("3 modelos", df_clean["ServiceModel"].nunique() == 3),
      ("Sin NaN en BilledCost", df_clean["BilledCost"].notna().all()),
  ]

  all_passed = True

  for description, passed in checks:
      status = "PASS" if passed else "FAIL"
      print(f"[{status}] {description}")
      if not passed:
          all_passed = False

  if all_passed:
      print("\nLABORATORIO COMPLETADO EXITOSAMENTE")
  else:
      print("\nHAY VERIFICACIONES FALLIDAS")
  ```

  > **Salida esperada:** Todas las verificaciones muestran `PASS`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el script completo para generar el reporte final.

  ```powershell
  python "$LabRoot\scripts\billing_analysis.py"
  ```

  > **Salida esperada:** El análisis finaliza sin errores y se crea `novatech_billing_summary_jan2024.xlsx`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Guarda el script y comprueba los entregables desde PowerShell.

  ```powershell
  Get-Item `
    "$LabRoot\data\raw\novatech_billing_raw.csv", `
    "$LabRoot\scripts\billing_analysis.py", `
    "$LabRoot\outputs\novatech_billing_summary_jan2024.xlsx" |
    Select-Object Name, Length
  ```

  > **Salida esperada:** Los tres archivos aparecen con tamaño mayor a 0 bytes.
  {: .lab-note .output .compact}

{% capture r3 %}{{ results[2] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r3 %}
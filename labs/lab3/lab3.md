---
layout: lab
title: "Demostración 3: Análisis de tendencias y anomalías de consumo"
permalink: /lab3/lab3/
images_base: /labs/lab3/img
duration: "25 minutos"
objective:
  - Analizar tendencias semanales y anomalías de gasto tecnológico mediante datos históricos normalizados en EUR y técnicas básicas de análisis FinOps.
prerequisites:
  - Equipo con Windows 11 y Visual Studio Code instalado.
  - Python 3.12 o superior disponible desde la terminal integrada de Visual Studio Code.
  - Paquetes pandas, numpy, matplotlib y openpyxl instalados.
  - Conocimiento básico de series de tiempo, variación Week-over-Week y FinOps Anomaly Management.
introduction:
  - En esta demostración analizarás cuatro meses de datos históricos de NovaTech SRL para identificar tendencias, variaciones Week-over-Week y anomalías de gasto. Trabajarás con un subconjunto de campos alineados con FOCUS 1.4, generarás visualizaciones y exportarás un reporte consolidado en Excel.
slug: lab3
lab_number: 3
final_result: >
  Al finalizar tendrás en el Escritorio un dataset histórico de cuatro meses en EUR, un reporte Excel con tendencias y anomalías, y dos gráficas PNG que resumen la evolución semanal del gasto y los eventos atípicos detectados.
notes:
  - Todos los archivos se almacenan dentro de la carpeta finops-course del Escritorio de Windows.
  - El dataset usa EUR como BillingCurrency y EffectiveCost como métrica principal para el análisis temporal.
  - El archivo utiliza un subconjunto de campos alineados con FOCUS 1.4; no representa una implementación completa de la especificación.
  - El z-score se usa como técnica estadística didáctica para detectar valores atípicos; FinOps no prescribe un algoritmo único de detección de anomalías.
references:
  - text: "FOCUS Specification"
    url: https://focus.finops.org/
  - text: "FinOps Anomaly Management"
    url: https://www.finops.org/framework/capabilities/anomaly-management/
  - text: "pandas Time Series"
    url: https://pandas.pydata.org/docs/user_guide/timeseries.html
  - text: "matplotlib Documentation"
    url: https://matplotlib.org/stable/
prev: /lab2/lab2/
next: /lab4/lab4/
---

---

<!-- Aquí comienzan las instrucciones paso a paso de la práctica -->

## 🔎 Tarea 1. Preparar y validar los datos históricos — 4 min

En esta tarea prepararás la estructura local de trabajo, descargarás el dataset histórico de NovaTech y comprobarás que contiene cuatro meses de información, tres categorías de servicio y costos expresados en EUR.

### Tarea 1.1. Preparar el directorio y descargar el dataset

Crearás la estructura necesaria dentro del Escritorio y descargarás el archivo histórico directamente desde la terminal integrada de Visual Studio Code.

- {% include step_label.html %} Abre Visual Studio Code y una terminal integrada de PowerShell.

  > **Nota:** Usa **Terminal → New Terminal** y confirma que el perfil activo sea PowerShell.
  {: .lab-note .info .compact}

  > **Salida esperada:** VS Code muestra una terminal PowerShell disponible para ejecutar comandos.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la carpeta del laboratorio y crea los directorios necesarios.

  > **Importante:** La variable `$LabRoot` evita depender del nombre específico del usuario de Windows.
  {: .lab-note .important .compact}

  ```powershell
  $LabRoot = Join-Path ([Environment]::GetFolderPath('Desktop')) 'finops-course'
  New-Item -ItemType Directory -Force -Path `
    "$LabRoot\data\raw", `
    "$LabRoot\data\processed", `
    "$LabRoot\outputs\lab3", `
    "$LabRoot\scripts" | Out-Null
  Set-Location $LabRoot
  ```

  > **Salida esperada:** Existen las carpetas `data\raw`, `data\processed`, `outputs\lab3` y `scripts`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la URL del dataset y descárgalo en `data\raw`.

  ```powershell
  $DatasetUrl = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_billing_4months_eur.csv"

  Invoke-WebRequest `
    -Uri $DatasetUrl `
    -OutFile "$LabRoot\data\raw\novatech_billing_4months_eur.csv"
  ```

  > **Salida esperada:** Se crea `data\raw\novatech_billing_4months_eur.csv` sin errores de descarga.
  {: .lab-note .output .compact}

### Tarea 1.2. Validar el dataset histórico

Verificarás que el archivo descargado tenga suficientes registros, cubra el período esperado y utilice EUR como moneda de facturación.

- {% include step_label.html %} Confirma que el archivo exista y tenga contenido.

  > **Advertencia:** No continúes si `Length` es 0 o si PowerShell indica que el archivo no existe.
  {: .lab-note .warning .compact}

  ```powershell
  Get-Item "$LabRoot\data\raw\novatech_billing_4months_eur.csv" |
    Select-Object Name, Length
  ```

  > **Salida esperada:** El archivo aparece con un tamaño mayor a 0 bytes.
  {: .lab-note .output .compact}

- {% include step_label.html %} Verifica cantidad de registros, período, categorías y moneda con Python.

  > **Nota:** Esta comprobación evita continuar con un dataset incompleto o diferente al esperado.
  {: .lab-note .info .compact}

  ```powershell
  python -c "import pandas as pd; p=r'$LabRoot\data\raw\novatech_billing_4months_eur.csv'; df=pd.read_csv(p); print('Registros:',len(df)); print('Periodo:',df['ChargePeriodStart'].min(),'a',df['ChargePeriodStart'].max()); print('Categorias:',sorted(df['ServiceCategory'].unique())); print('Monedas:',df['BillingCurrency'].unique().tolist())"
  ```

  > **Salida esperada:** Se muestran aproximadamente 3,700 registros, período `2023-10-01` a `2024-01-31`, categorías IaaS/PaaS/SaaS y moneda EUR.
  {: .lab-note .output .compact}

{% assign results = site.data.task-results[page.slug].results %}
{% capture r1 %}{{ results[0] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r1 %}

---

## 📈 Tarea 2. Analizar tendencias y detectar anomalías — 12 min

En esta tarea transformarás los registros diarios en series semanales, calcularás variaciones Week-over-Week y aplicarás z-score para identificar picos o caídas que requieran investigación FinOps.

### Tarea 2.1. Construir la serie semanal y calcular WoW

Crearás un script Python que agrupe el EffectiveCost por semana y servicio, calcule variaciones WoW y determine los servicios con mayor crecimiento reciente.

- {% include step_label.html %} Crea el archivo `scripts\trend_analysis.py` desde Visual Studio Code.

  > **Nota:** Usa el explorador de VS Code para crear el archivo dentro de la carpeta `scripts`.
  {: .lab-note .info .compact}

  > **Salida esperada:** Existe el archivo `scripts\trend_analysis.py`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega el código para cargar y preparar la serie temporal semanal.

  > **Importante:** `EffectiveCost` se usa como métrica principal de costo para el análisis temporal.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\trend_analysis.py`, pega este bloque desde la **primera línea del archivo**, ya que corresponde al código inicial del script.

  ```python
  from pathlib import Path
  import pandas as pd
  import numpy as np
  import matplotlib.pyplot as plt

  ROOT = Path.home() / "Desktop" / "finops-course"
  RAW = ROOT / "data" / "raw" / "novatech_billing_4months_eur.csv"
  PROCESSED = ROOT / "data" / "processed"
  OUTPUTS = ROOT / "outputs" / "lab3"

  PROCESSED.mkdir(parents=True, exist_ok=True)
  OUTPUTS.mkdir(parents=True, exist_ok=True)

  df = pd.read_csv(RAW, parse_dates=["ChargePeriodStart"])

  df["Week"] = (
      df["ChargePeriodStart"]
      .dt.to_period("W-MON")
      .apply(lambda r: r.start_time)
  )

  weekly_service = (
      df.groupby(["Week", "ServiceName", "ServiceCategory"], as_index=False)
        .agg(
            EffectiveCost=("EffectiveCost", "sum"),
            Records=("RecordId", "count")
        )
  )

  pivot_services = weekly_service.pivot_table(
      index="Week",
      columns="ServiceName",
      values="EffectiveCost",
      aggfunc="sum",
      fill_value=0
  )

  print(f"Registros cargados: {len(df):,}")
  print(f"Semanas analizadas: {pivot_services.shape[0]}")
  print(f"Servicios analizados: {pivot_services.shape[1]}")
  ```

  > **Salida esperada:** El script queda preparado para generar una tabla semanal por servicio.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega el cálculo de variación Week-over-Week y crecimiento reciente.

  > **Nota:** Se compara el promedio de las primeras cuatro semanas con el de las últimas cuatro para reducir el efecto de semanas parciales.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\trend_analysis.py`, agrega este bloque **debajo de** `print(f"Servicios analizados: {pivot_services.shape[1]}")`.

  ```python
  wow_change = pivot_services.pct_change() * 100
  wow_change = wow_change.replace([np.inf, -np.inf], np.nan)

  first4 = pivot_services.head(4).mean()
  last4 = pivot_services.tail(4).mean()

  growth = ((last4 - first4) / first4.replace(0, np.nan) * 100)
  growth = growth.dropna().sort_values(ascending=False)

  print("\nTop 5 servicios por crecimiento:")
  print(growth.head(5).round(1).to_string())
  ```

  > **Salida esperada:** Se obtiene un ranking de cinco servicios con mayor crecimiento entre el inicio y el final del período.
  {: .lab-note .output .compact}

### Tarea 2.2. Detectar anomalías con z-score

Aplicarás un z-score por servicio para identificar semanas con comportamiento atípico y distinguir picos de gasto de caídas significativas.

- {% include step_label.html %} Agrega una función para calcular z-score por servicio.

  > **Nota:** El z-score mide cuántas desviaciones estándar se aleja un valor respecto a la media histórica del servicio.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\trend_analysis.py`, agrega este bloque **debajo de** `print(growth.head(5).round(1).to_string())`.

  ```python
  def calculate_zscore(series):
      std = series.std()
      if std == 0 or pd.isna(std):
          return pd.Series(0, index=series.index)
      return (series - series.mean()) / std

  zscores = pivot_services.apply(calculate_zscore)
  ```

  > **Salida esperada:** Cada servicio dispone de una serie semanal de z-scores.
  {: .lab-note .output .compact}

- {% include step_label.html %} Identifica semanas con `|z-score| > 2.0`.

  > **Importante:** Un valor atípico indica que debe investigarse; por sí solo no demuestra desperdicio ni error.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\trend_analysis.py`, agrega este bloque **debajo de** `zscores = pivot_services.apply(calculate_zscore)`.

  ```python
  threshold = 2.0
  anomalies = []

  for service in zscores.columns:
      for week in zscores.index:
          z = zscores.loc[week, service]

          if abs(z) > threshold:
              cost = pivot_services.loc[week, service]

              anomalies.append({
                  "Week": week,
                  "ServiceName": service,
                  "ZScore": round(z, 2),
                  "EffectiveCostEUR": round(cost, 2),
                  "Type": "PICO" if z > 0 else "CAIDA"
              })

  df_anomalies = pd.DataFrame(anomalies)

  if not df_anomalies.empty:
      df_anomalies = df_anomalies.sort_values(
          "ZScore",
          key=lambda s: s.abs(),
          ascending=False
      )

  print(f"\nAnomalías detectadas: {len(df_anomalies)}")
  print(df_anomalies.head(10).to_string(index=False))
  ```

  > **Salida esperada:** Se detectan varias anomalías de tipo `PICO` o `CAIDA` con z-score absoluto superior a 2.0.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el script y revisa los resultados principales.

  > **Advertencia:** Si no se detectan anomalías, verifica que estés usando el dataset correcto y que el archivo no haya sido modificado.
  {: .lab-note .warning .compact}

  ```powershell
  python "$LabRoot\scripts\trend_analysis.py"
  ```

  > **Salida esperada:** Se muestran registros, semanas, servicios, top 5 de crecimiento y anomalías detectadas.
  {: .lab-note .output .compact}

{% capture r2 %}{{ results[1] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r2 %}

---

## 🚀 Tarea 3. Visualizar y exportar los resultados — 9 min

En esta tarea generarás dos visualizaciones independientes y exportarás un reporte Excel consolidado para comunicar las tendencias y anomalías detectadas.

### Tarea 3.1. Generar visualizaciones de tendencia y anomalías

Agregarás dos gráficas al script: una para los servicios con mayor crecimiento y otra para representar las anomalías detectadas por servicio.

- {% include step_label.html %} Agrega una gráfica de líneas para los cinco servicios con mayor crecimiento.

  > **Nota:** La gráfica permite comparar la evolución semanal sin asumir que crecimiento equivale automáticamente a ineficiencia.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\trend_analysis.py`, agrega este bloque **debajo de** `print(df_anomalies.head(10).to_string(index=False))`.

  ```python
  top5_services = growth.head(5).index.tolist()

  plt.figure(figsize=(12, 6))

  for service in top5_services:
      plt.plot(
          pivot_services.index,
          pivot_services[service],
          marker="o",
          linewidth=2,
          label=service
      )

  plt.xlabel("Semana")
  plt.ylabel("EffectiveCost semanal (EUR)")
  plt.legend()
  plt.grid(alpha=0.3)
  plt.xticks(rotation=45)
  plt.tight_layout()

  trend_png = OUTPUTS / "trend_top5.png"
  plt.savefig(trend_png, dpi=150, bbox_inches="tight")
  plt.close()

  print(f"Gráfica creada: {trend_png}")
  ```

  > **Salida esperada:** Se crea `outputs\lab3\trend_top5.png`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega una gráfica de dispersión con las anomalías detectadas.

  > **Importante:** La visualización muestra eventos que requieren investigación, no conclusiones automáticas sobre desperdicio.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\trend_analysis.py`, agrega este bloque **debajo de** `print(f"Gráfica creada: {trend_png}")`.

  ```python
  plt.figure(figsize=(12, 6))

  if not df_anomalies.empty:
      anomaly_plot = df_anomalies.copy()

      plt.scatter(
          anomaly_plot["Week"],
          anomaly_plot["EffectiveCostEUR"],
          s=70
      )

      for _, row in anomaly_plot.iterrows():
          plt.annotate(
              row["ServiceName"],
              (row["Week"], row["EffectiveCostEUR"]),
              xytext=(5, 5),
              textcoords="offset points",
              fontsize=8
          )

  plt.xlabel("Semana")
  plt.ylabel("EffectiveCost anómalo (EUR)")
  plt.grid(alpha=0.3)
  plt.xticks(rotation=45)
  plt.tight_layout()

  anomaly_png = OUTPUTS / "anomalies.png"
  plt.savefig(anomaly_png, dpi=150, bbox_inches="tight")
  plt.close()

  print(f"Gráfica creada: {anomaly_png}")
  ```

  > **Salida esperada:** Se crea `outputs\lab3\anomalies.png` con los eventos atípicos identificados.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta nuevamente el script para generar ambas gráficas.

  > **Nota:** Los archivos PNG se sobrescriben si ya existían de una ejecución anterior.
  {: .lab-note .info .compact}

  ```powershell
  python "$LabRoot\scripts\trend_analysis.py"
  ```

  > **Salida esperada:** La terminal confirma la creación de `trend_top5.png` y `anomalies.png`.
  {: .lab-note .output .compact}

### Tarea 3.2. Exportar y validar el reporte consolidado

Exportarás las series semanales, variaciones WoW, anomalías y ranking de crecimiento a un archivo Excel reutilizable.

- {% include step_label.html %} Agrega al final del script la exportación de `novatech_trend_report.xlsx`.

  > **Nota:** El reporte consolida los principales artefactos analíticos en un único archivo.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\trend_analysis.py`, agrega este bloque **al final del archivo**, **debajo de** `print(f"Gráfica creada: {anomaly_png}")`.

  ```python
  report_xlsx = PROCESSED / "novatech_trend_report.xlsx"

  growth_df = pd.DataFrame({
      "ServiceName": growth.index,
      "GrowthPct": growth.values,
      "First4WeeksAvgEUR": [first4[s] for s in growth.index],
      "Last4WeeksAvgEUR": [last4[s] for s in growth.index]
  })

  summary = pd.DataFrame({
      "Metric": [
          "PeriodStart",
          "PeriodEnd",
          "TotalEffectiveCostEUR",
          "AverageWeeklyCostEUR",
          "ServicesAnalyzed",
          "AnomaliesDetected"
      ],
      "Value": [
          df["ChargePeriodStart"].min().date().isoformat(),
          df["ChargePeriodStart"].max().date().isoformat(),
          round(df["EffectiveCost"].sum(), 2),
          round(pivot_services.sum(axis=1).mean(), 2),
          pivot_services.shape[1],
          len(df_anomalies)
      ]
  })

  with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
      summary.to_excel(writer, sheet_name="summary", index=False)
      pivot_services.to_excel(writer, sheet_name="weekly_by_service")
      wow_change.to_excel(writer, sheet_name="wow_change_pct")

      if not df_anomalies.empty:
          df_anomalies.to_excel(writer, sheet_name="anomalies", index=False)

      growth_df.to_excel(writer, sheet_name="growth", index=False)

  print(f"Reporte creado: {report_xlsx}")
  ```

  > **Salida esperada:** Se crea `data\processed\novatech_trend_report.xlsx` con las hojas de resumen, tendencias, WoW, anomalías y crecimiento.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el script y valida los tres artefactos finales.

  > **Importante:** Los tres archivos deben existir y tener un tamaño mayor a 0 bytes.
  {: .lab-note .important .compact}

  ```powershell
  python "$LabRoot\scripts\trend_analysis.py"
  ```
  ```powershell
  Get-Item `
    "$LabRoot\data\processed\novatech_trend_report.xlsx", `
    "$LabRoot\outputs\lab3\trend_top5.png", `
    "$LabRoot\outputs\lab3\anomalies.png" |
    Select-Object Name, Length
  ```

  > **Salida esperada:** PowerShell muestra el Excel y las dos imágenes con tamaños mayores a 0 bytes.
  {: .lab-note .output .compact}

{% capture r3 %}{{ results[2] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r3 %}
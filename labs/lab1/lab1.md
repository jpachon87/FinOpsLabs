---
layout: lab
title: "Demostración 1: Diagnóstico inicial de problemas FinOps"
permalink: /lab1/lab1/
images_base: /labs/lab1/img
duration: "15 minutos"
objective:
  - Diagnosticar síntomas de gestión tecnológica y financiera mediante Domains, Capabilities, Phases y madurez del FinOps Framework 2026.
prerequisites:
  - Equipo con Windows 11 y Visual Studio Code instalado.
  - PowerShell disponible desde la terminal integrada de Visual Studio Code.
  - Conexión a Internet para descargar el dataset del laboratorio.
  - Conocimiento básico de Domains, Capabilities, Phases y Maturity Model de FinOps.
introduction:
  - En esta demostración analizarás el caso ficticio de NovaTech SRL, una organización con gasto tecnológico distribuido entre AWS, Azure y GCP. Usarás Visual Studio Code y PowerShell para revisar síntomas, clasificarlos con el FinOps Framework 2026 y obtener un diagnóstico preliminar de prioridades y madurez.
slug: lab1
lab_number: 1
final_result: >
  Al finalizar tendrás en el Escritorio una estructura de trabajo con el dataset original, un archivo CSV con el diagnóstico de 10 síntomas y un resumen en Markdown con los principales hallazgos FinOps expresados en euros.
notes:
  - Todos los archivos de la práctica se almacenan en la carpeta finops-course del Escritorio de Windows.
  - Los importes del caso están expresados en euros y la columna impacto_estimado_mensual_eur ya está normalizada a valor mensual.
  - La fase indicada es la fase predominante para abordar cada síntoma; Inform, Optimize y Operate forman un ciclo iterativo.
  - La madurez Crawl, Walk y Run no se deriva directamente de las Phases y debe interpretarse por Capability, Domain o actividad evaluada.
references:
  - text: "FinOps Framework"
    url: https://www.finops.org/framework/
  - text: "FinOps Domains"
    url: https://www.finops.org/framework/domains/
  - text: "FinOps Capabilities"
    url: https://www.finops.org/framework/capabilities/
  - text: "FinOps Phases"
    url: https://www.finops.org/framework/phases/
  - text: "FinOps Maturity Model"
    url: https://www.finops.org/framework/maturity-model/
prev: /
next: /lab2/lab2/
---

---

<!-- Aquí comienzan las instrucciones paso a paso de la práctica -->

## 🔎 Tarea 1. Preparar y revisar el escenario de NovaTech — 4 min

En esta tarea prepararás la estructura local del laboratorio en el Escritorio, incorporarás el dataset de NovaTech y comprobarás que contiene los 10 síntomas necesarios para realizar el diagnóstico FinOps.

### Tarea 1.1. Preparar el directorio de trabajo

Crearás la estructura mínima de carpetas en el Escritorio y descargarás el archivo de entrada directamente en `data\raw` desde VS Code.

- {% include step_label.html %} Abre Visual Studio Code y una terminal integrada de PowerShell.

  > **Nota:** Usa **Terminal → New Terminal** y confirma que el perfil activo sea PowerShell.
  {: .lab-note .info .compact}

  > **Salida esperada:** VS Code muestra una terminal PowerShell lista para ejecutar comandos.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea la estructura del laboratorio dentro del Escritorio de Windows.

  > **Importante:** La variable `$LabRoot` evita depender del nombre específico del usuario de Windows.
  {: .lab-note .important .compact}

  ```powershell
  $LabRoot = Join-Path ([Environment]::GetFolderPath('Desktop')) 'finops-course'
  New-Item -ItemType Directory -Force -Path "$LabRoot\data\raw", "$LabRoot\data\processed", "$LabRoot\outputs" | Out-Null
  Set-Location $LabRoot
  ```

  > **Salida esperada:** La terminal queda ubicada en `Desktop\finops-course` y existen las carpetas `data\raw`, `data\processed` y `outputs`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Descarga `novatech_symptoms_eur.csv` directamente en la carpeta `data\raw`.

  > **Nota:** El archivo se obtiene desde el repositorio de recursos del curso y se guarda directamente dentro del laboratorio.
  {: .lab-note .info .compact}

  ```powershell
  Invoke-WebRequest `
    -Uri "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_symptoms_eur.csv" `
    -OutFile "$LabRoot\data\raw\novatech_symptoms_eur.csv"
  ```

  > **Salida esperada:** El comando finaliza sin errores y crea `data\raw\novatech_symptoms_eur.csv`.
  {: .lab-note .output .compact}

### Tarea 1.2. Verificar y explorar los síntomas

Comprobarás la estructura del CSV y revisarás rápidamente los síntomas para identificar las áreas donde existen problemas de uso, costo, valor y práctica FinOps.

- {% include step_label.html %} Verifica que el archivo de entrada exista en la ruta esperada.

  > **Advertencia:** No continúes si el comando devuelve `False`; corrige primero la ubicación del archivo.
  {: .lab-note .warning .compact}

  ```powershell
  Test-Path "$LabRoot\data\raw\novatech_symptoms_eur.csv"
  ```

  > **Salida esperada:** PowerShell devuelve `True`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Visualiza el encabezado y los primeros registros del dataset.

  > **Nota:** El dataset usa `impacto_estimado_mensual_eur`, por lo que los importes son comparables entre síntomas.
  {: .lab-note .info .compact}

  ```powershell
  Import-Csv "$LabRoot\data\raw\novatech_symptoms_eur.csv" |
    Select-Object -First 4 |
    Format-Table id, area_afectada, frecuencia, impacto_estimado_mensual_eur -AutoSize
  ```

  > **Salida esperada:** Se muestran los primeros cuatro síntomas con identificador, área, frecuencia e impacto mensual en EUR.
  {: .lab-note .output .compact}

- {% include step_label.html %} Confirma que el dataset contenga exactamente 10 síntomas.

  > **Importante:** Los identificadores deben abarcar de `S01` a `S10`.
  {: .lab-note .important .compact}

  ```powershell
  (Import-Csv "$LabRoot\data\raw\novatech_symptoms_eur.csv").Count
  ```

  > **Salida esperada:** PowerShell devuelve `10`.
  {: .lab-note .output .compact}

{% assign results = site.data.task-results[page.slug].results %}
{% capture r1 %}{{ results[0] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r1 %}

---

## ☁️ Tarea 2. Clasificar los síntomas con FinOps Framework 2026 — 7 min

En esta tarea crearás el diagnóstico estructurado de NovaTech y revisarás cómo cada síntoma se relaciona con un Domain, una Capability y una fase predominante para su tratamiento.

### Tarea 2.1. Crear el archivo de diagnóstico FinOps

Generarás un CSV con el mapeo de los 10 síntomas utilizando nombres vigentes de Domains y Capabilities del FinOps Framework 2026.

- {% include step_label.html %} Crea `diagnostico_finops_novatech.csv` con la clasificación completa del caso.

  > **Importante:** `tipo_iniciativa_laboratorio` es una clasificación didáctica de esta práctica y no una taxonomía oficial del Framework.
  {: .lab-note .important .compact}

  ```powershell
  @'
  id,domain,capability,fase_predominante,tipo_iniciativa_laboratorio,causa_raiz,prioridad
  S01,"Understand Usage & Cost","Anomaly Management","Inform","Visibilidad y control","Ausencia de monitoreo continuo y gestión de anomalías","Alta"
  S02,"Optimize Usage & Cost","Usage Optimization","Optimize","Optimización del gasto","Recursos activos fuera de horario sin necesidad de negocio","Media"
  S03,"Understand Usage & Cost","Allocation","Inform","Visibilidad y control","Etiquetado incompleto y estructuras de asignación inconsistentes","Alta"
  S04,"Optimize Usage & Cost","Usage Optimization","Optimize","Optimización del gasto","Recursos de alto costo sin políticas de apagado","Media"
  S05,"Quantify Business Value","Budgeting","Inform","Planificación y valor","Presupuestos sin umbrales ni alertas automáticas","Alta"
  S06,"Optimize Usage & Cost","Usage Optimization","Optimize","Optimización del gasto","Capacidad de staging sobredimensionada respecto a su demanda","Media"
  S07,"Manage the FinOps Practice","FinOps Practice Operations","Operate","Gestión de la práctica","Responsabilidades FinOps y accountability no definidos","Alta"
  S08,"Optimize Usage & Cost","Rate Optimization","Optimize","Optimización del gasto","Compromisos sin revisión periódica de utilización","Media"
  S09,"Manage the FinOps Practice","FinOps Education & Enablement","Inform","Gestión de la práctica","Ingeniería sin visibilidad ni formación sobre costos","Alta"
  S10,"Quantify Business Value","Forecasting","Inform","Planificación y valor","Forecast manual y poco frecuente sin señales recientes","Media"
  '@ | Set-Content "$LabRoot\data\processed\diagnostico_finops_novatech.csv" -Encoding UTF8
  ```

  > **Salida esperada:** Se crea `data\processed\diagnostico_finops_novatech.csv` con 10 filas de diagnóstico.
  {: .lab-note .output .compact}

- {% include step_label.html %} Comprueba que los cuatro Domains de FinOps estén representados.

  > **Nota:** Los Domains describen resultados de negocio y agrupan Capabilities relacionadas.
  {: .lab-note .info .compact}

  ```powershell
  Import-Csv "$LabRoot\data\processed\diagnostico_finops_novatech.csv" |
    Group-Object domain |
    Select-Object Name, Count |
    Format-Table -AutoSize
  ```

  > **Salida esperada:** Aparecen `Understand Usage & Cost`, `Quantify Business Value`, `Optimize Usage & Cost` y `Manage the FinOps Practice`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Revisa tres casos representativos para relacionar síntoma, Domain y Capability.

  > **Nota:** S01 ejemplifica anomalías, S03 asignación de costos y S07 operación de la práctica FinOps.
  {: .lab-note .info .compact}

  ```powershell
  Import-Csv "$LabRoot\data\processed\diagnostico_finops_novatech.csv" |
    Where-Object id -in 'S01','S03','S07' |
    Format-Table id, domain, capability, prioridad -Wrap
  ```

  > **Salida esperada:** Los tres síntomas muestran clasificaciones distintas y coherentes con su causa principal.
  {: .lab-note .output .compact}

### Tarea 2.2. Analizar fase predominante, causa y prioridad

Revisarás cómo la fase predominante ayuda a orientar la respuesta sin confundirla con el nivel de madurez Crawl, Walk o Run.

- {% include step_label.html %} Agrupa los síntomas por fase predominante.

  > **Importante:** Inform, Optimize y Operate son fases iterativas; aquí se usa la fase que mejor representa la acción inicial.
  {: .lab-note .important .compact}

  ```powershell
  Import-Csv "$LabRoot\data\processed\diagnostico_finops_novatech.csv" |
    Group-Object fase_predominante |
    Select-Object Name, Count |
    Format-Table -AutoSize
  ```

  > **Salida esperada:** `Inform = 5`, `Optimize = 4` y `Operate = 1`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Identifica los síntomas considerados de prioridad alta.

  > **Nota:** La prioridad del laboratorio combina impacto organizacional, visibilidad y urgencia de actuación.
  {: .lab-note .info .compact}

  ```powershell
  Import-Csv "$LabRoot\data\processed\diagnostico_finops_novatech.csv" |
    Where-Object prioridad -eq 'Alta' |
    Select-Object id, capability, causa_raiz |
    Format-Table -Wrap
  ```

  > **Salida esperada:** Se muestran cinco síntomas de prioridad alta: S01, S03, S05, S07 y S09.
  {: .lab-note .output .compact}

- {% include step_label.html %} Contrasta los síntomas de optimización con los de gestión de la práctica.

  > **Nota:** FinOps combina decisiones tecnológicas, financieras y organizacionales; no se limita a reducir costos.
  {: .lab-note .info .compact}

  ```powershell
  Import-Csv "$LabRoot\data\processed\diagnostico_finops_novatech.csv" |
    Where-Object { $_.domain -in 'Optimize Usage & Cost','Manage the FinOps Practice' } |
    Format-Table id, domain, capability, fase_predominante -AutoSize
  ```

  > **Salida esperada:** Se distinguen problemas de optimización de uso o tarifas y problemas de operación o enablement de la práctica.
  {: .lab-note .output .compact}

{% capture r2 %}{{ results[1] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r2 %}

---

## 🚀 Tarea 3. Interpretar y validar el diagnóstico — 4 min

En esta tarea resumirás los hallazgos cuantitativos del caso y generarás una conclusión breve sobre las áreas que muestran menor madurez FinOps, sin derivarla directamente de las Phases.

### Tarea 3.1. Obtener los indicadores principales

Calcularás el impacto mensual cuantificado y revisarás la distribución de prioridades para sustentar el resumen ejecutivo de NovaTech.

- {% include step_label.html %} Calcula el impacto financiero mensual cuantificado en el dataset.

  > **Importante:** El cálculo solo suma síntomas con impacto monetario estimado y no asigna valor artificial a los registros con cero.
  {: .lab-note .important .compact}

  ```powershell
  $TotalEur = (Import-Csv "$LabRoot\data\raw\novatech_symptoms_eur.csv" |
    Measure-Object -Property impacto_estimado_mensual_eur -Sum).Sum
  "Impacto mensual cuantificado: €{0:N0}" -f $TotalEur
  ```

  > **Salida esperada:** `Impacto mensual cuantificado: €120,500`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Revisa la distribución de prioridades del diagnóstico.

  > **Nota:** Esta vista permite separar los problemas que requieren atención inmediata de los que pueden programarse.
  {: .lab-note .info .compact}

  ```powershell
  Import-Csv "$LabRoot\data\processed\diagnostico_finops_novatech.csv" |
    Group-Object prioridad |
    Select-Object Name, Count |
    Format-Table -AutoSize
  ```

  > **Salida esperada:** Se muestran `Alta = 5` y `Media = 5`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Genera un resumen por Domain para observar dónde se concentran los síntomas.

  > **Nota:** El número de síntomas por Domain es un indicador descriptivo y no equivale por sí solo a una medición de madurez.
  {: .lab-note .info .compact}

  ```powershell
  Import-Csv "$LabRoot\data\processed\diagnostico_finops_novatech.csv" |
    Group-Object domain |
    Sort-Object Count -Descending |
    Select-Object Name, Count |
    Format-Table -AutoSize
  ```

  > **Salida esperada:** Los 10 síntomas quedan distribuidos entre los cuatro Domains del FinOps Framework.
  {: .lab-note .output .compact}

### Tarea 3.2. Crear y validar el resumen ejecutivo

Crearás un archivo Markdown con las conclusiones principales y comprobarás que los tres artefactos del laboratorio estén disponibles.

- {% include step_label.html %} Genera `resumen_finops_novatech.md` con la conclusión diagnóstica.

  > **Importante:** La referencia a Crawl es preliminar y se aplica a varias Capabilities observadas, no a la organización por contar fases.
  {: .lab-note .important .compact}

  ```powershell
  @"
  # Resumen diagnóstico FinOps - NovaTech SRL

  - Síntomas analizados: 10
  - Prioridad alta: 5
  - Prioridad media: 5
  - Fase predominante Inform: 5
  - Fase predominante Optimize: 4
  - Fase predominante Operate: 1
  - Impacto mensual cuantificado: €120,500

  ## Conclusión

  Los síntomas sugieren una madurez preliminar predominantemente Crawl en varias
  Capabilities, especialmente Anomaly Management, Allocation, Budgeting,
  Forecasting, Usage Optimization y FinOps Practice Operations. Esta conclusión
  es diagnóstica y no constituye una evaluación formal de madurez organizacional.
  NovaTech debe mejorar visibilidad, accountability, planificación y optimización
  de uso antes de escalar automatización y prácticas operativas más avanzadas.
  "@ | Set-Content "$LabRoot\outputs\resumen_finops_novatech.md" -Encoding UTF8
  ```

  > **Salida esperada:** Se crea `outputs\resumen_finops_novatech.md` con métricas y una conclusión alineada al Maturity Model.
  {: .lab-note .output .compact}

- {% include step_label.html %} Verifica que los tres archivos principales de la demostración existan.

  > **Nota:** Esta validación confirma que el dataset, el diagnóstico y el resumen quedaron guardados en el Escritorio.
  {: .lab-note .info .compact}

  ```powershell
  Get-Item `
    "$LabRoot\data\raw\novatech_symptoms_eur.csv", `
    "$LabRoot\data\processed\diagnostico_finops_novatech.csv", `
    "$LabRoot\outputs\resumen_finops_novatech.md" |
    Select-Object Name, Length
  ```

  > **Salida esperada:** PowerShell muestra los tres archivos con un tamaño mayor a 0 bytes.
  {: .lab-note .output .compact}

{% capture r3 %}{{ results[2] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r3 %}
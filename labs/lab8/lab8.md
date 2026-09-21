---
layout: lab
title: "Demostración 8: Arquitectura de datos FinOps con FOCUS"
permalink: /lab8/lab8/
images_base: /labs/lab8/img
duration: "30 minutos"
objective:
  - Implementar una arquitectura de datos FinOps multicloud que ingeste datos heterogéneos, los normalice a un modelo alineado con un subconjunto de FOCUS 1.4, aplique validaciones de calidad y publique métricas en Grafana.
prerequisites:
  - Equipo con Windows 11, Visual Studio Code y Docker Desktop instalados.
  - Docker Compose disponible desde PowerShell.
  - Python 3.12 o superior disponible desde la terminal integrada de Visual Studio Code.
  - Paquetes pandas, SQLAlchemy, psycopg2-binary, dbt-core y dbt-postgres instalados.
  - Conocimiento básico de SQL, Docker, dbt y conceptos FOCUS.
introduction:
  - En esta demostración trabajarás con tres archivos de billing multicloud de NovaTech SRL con esquemas heterogéneos. Los cargarás a PostgreSQL, los normalizarás mediante dbt a un modelo FOCUS 1.4-aligned subset, validarás calidad de datos y visualizarás métricas FinOps en Grafana.
slug: lab8
lab_number: 8
final_result: >
  Al finalizar tendrás PostgreSQL con tres tablas staging y un mart unificado alineado con un subconjunto de FOCUS 1.4, validaciones dbt, un dashboard Grafana y documentación breve de gobierno de datos.
notes:
  - Todos los archivos se almacenan dentro de la carpeta finops-course del Escritorio de Windows.
  - Todos los importes monetarios se expresan en EUR.
  - Esta práctica implementa un subconjunto alineado con FOCUS 1.4 y no una implementación completa de todas las columnas obligatorias de la especificación.
  - Los nombres SQL usan snake_case para evitar problemas de quoting en PostgreSQL.
  - Los objetivos de freshness y calidad documentados son políticas internas de NovaTech.
references:
  - text: "FOCUS Specification"
    url: https://focus.finops.org/
  - text: "FinOps Data Ingestion"
    url: https://www.finops.org/framework/capabilities/data-ingestion/
  - text: "dbt Documentation"
    url: https://docs.getdbt.com/
  - text: "Grafana PostgreSQL Data Source"
    url: https://grafana.com/docs/grafana/latest/datasources/postgres/
prev: /lab7/lab7/
next: /lab9/lab9/
---

---

<!-- Aquí comienzan las instrucciones paso a paso de la práctica -->

## 🧱 Tarea 1. Levantar la arquitectura e ingresar los datos — 7 min

En esta tarea prepararás PostgreSQL y Grafana con Docker Compose, descargarás los tres archivos de billing y los cargarás en tablas staging.

### Tarea 1.1. Preparar Docker y descargar los datasets

- {% include step_label.html %} Abre Visual Studio Code y una terminal integrada de PowerShell.

  > **Nota:** Usa **Terminal → New Terminal** y confirma que el perfil activo sea PowerShell.
  {: .lab-note .info .compact}

  > **Salida esperada:** VS Code muestra una terminal PowerShell lista para ejecutar comandos.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la carpeta del laboratorio y crea la estructura de trabajo.

  > **Importante:** La variable `$LabRoot` permite trabajar sin depender del nombre específico del usuario de Windows.
  {: .lab-note .important .compact}

  ```powershell
  $LabRoot = Join-Path ([Environment]::GetFolderPath('Desktop')) 'finops-course'

  New-Item -ItemType Directory -Force -Path `
    "$LabRoot\data\raw", `
    "$LabRoot\docker\lab8", `
    "$LabRoot\dbt\lab8\models\staging", `
    "$LabRoot\dbt\lab8\models\marts", `
    "$LabRoot\scripts", `
    "$LabRoot\outputs\lab8", `
    "$LabRoot\docs\lab8" | Out-Null

  Set-Location $LabRoot
  ```

  > **Salida esperada:** Se crean las carpetas de datos, Docker, dbt, scripts, outputs y documentación.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define las tres URL de descarga y descarga los CSV.

  ```powershell
  $AwsUrl   = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_aws_billing_raw_eur.csv"
  $AzureUrl = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_azure_billing_raw_eur.csv"
  $GcpUrl   = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_gcp_billing_raw_eur.csv"
  ```
  ```powershell
  Invoke-WebRequest `
    -Uri $AwsUrl `
    -OutFile "$LabRoot\data\raw\novatech_aws_billing_raw_eur.csv"

  Invoke-WebRequest `
    -Uri $AzureUrl `
    -OutFile "$LabRoot\data\raw\novatech_azure_billing_raw_eur.csv"

  Invoke-WebRequest `
    -Uri $GcpUrl `
    -OutFile "$LabRoot\data\raw\novatech_gcp_billing_raw_eur.csv"
  ```

  > **Salida esperada:** Se descargan los tres archivos multicloud en `data\raw`.
  {: .lab-note .output .compact}

### Tarea 1.2. Levantar PostgreSQL/Grafana e ingresar staging

- {% include step_label.html %} Crea `docker\lab8\docker-compose.yml`.

  > **Nota:** PostgreSQL será el data warehouse del laboratorio y Grafana la capa de visualización.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `docker\lab8\docker-compose.yml`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```yaml
  services:
    postgres:
      image: postgres:16
      container_name: finops-postgres
      environment:
        POSTGRES_DB: finops_dw
        POSTGRES_USER: finops_admin
        POSTGRES_PASSWORD: finops2026
      ports:
        - "5432:5432"
      healthcheck:
        test: ["CMD-SHELL", "pg_isready -U finops_admin -d finops_dw"]
        interval: 5s
        timeout: 5s
        retries: 10

    grafana:
      image: grafana/grafana:latest
      container_name: finops-grafana
      environment:
        GF_SECURITY_ADMIN_USER: admin
        GF_SECURITY_ADMIN_PASSWORD: finops2026
        GF_USERS_ALLOW_SIGN_UP: "false"
      ports:
        - "3000:3000"
      depends_on:
        postgres:
          condition: service_healthy
  ```

  > **Salida esperada:** El archivo Docker Compose define PostgreSQL y Grafana.
  {: .lab-note .output .compact}

- {% include step_label.html %} Levanta los servicios y valida su estado.

  > **Importante:** Docker Desktop debe estar ejecutándose antes de continuar.
  {: .lab-note .important .compact}

  ```powershell
  docker compose -f "$LabRoot\docker\lab8\docker-compose.yml" up -d
  docker compose -f "$LabRoot\docker\lab8\docker-compose.yml" ps
  ```

  > **Salida esperada:** `finops-postgres` aparece healthy y `finops-grafana` en ejecución.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea `scripts\ingest_multicloud.py` para cargar las tres fuentes en staging.

  > **Nota:** Las fuentes conservan sus esquemas heterogéneos en staging; la normalización se realizará después con dbt.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\ingest_multicloud.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  from pathlib import Path
  import pandas as pd
  from sqlalchemy import create_engine, text

  ROOT = Path.home() / "Desktop" / "finops-course"
  RAW = ROOT / "data" / "raw"

  engine = create_engine(
      "postgresql+psycopg2://finops_admin:finops2026@localhost:5432/finops_dw"
  )

  with engine.begin() as conn:
      conn.execute(text("CREATE SCHEMA IF NOT EXISTS staging"))
      conn.execute(text("CREATE SCHEMA IF NOT EXISTS marts"))

  sources = {
      "aws_costs": RAW / "novatech_aws_billing_raw_eur.csv",
      "azure_costs": RAW / "novatech_azure_billing_raw_eur.csv",
      "gcp_costs": RAW / "novatech_gcp_billing_raw_eur.csv"
  }

  for table, path in sources.items():
      df = pd.read_csv(path)
      df.to_sql(
          table,
          engine,
          schema="staging",
          if_exists="replace",
          index=False
      )
      print(f"{table}: {len(df):,} registros")
  ```

  > **Salida esperada:** Se cargan 1,200 registros AWS, 1,100 Azure y 1,000 GCP.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el script de ingestión para crear las tablas staging.

  > **Importante:** Este paso debe completarse antes de ejecutar dbt, porque los modelos utilizan `staging.aws_costs`, `staging.azure_costs` y `staging.gcp_costs`.
  {: .lab-note .important .compact}

  ```powershell
  python "$LabRoot\scripts\ingest_multicloud.py"
  ```

  > **Salida esperada:** La terminal confirma `aws_costs: 1,200 registros`, `azure_costs: 1,100 registros` y `gcp_costs: 1,000 registros`.
  {: .lab-note .output .compact}

{% assign results = site.data.task-results[page.slug].results %}
{% capture r1 %}{{ results[0] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r1 %}

---

## 🔄 Tarea 2. Normalizar a un modelo FOCUS 1.4-aligned — 10 min

En esta tarea configurarás dbt y transformarás las tres fuentes heterogéneas a un esquema unificado de análisis.

### Tarea 2.1. Configurar dbt

- {% include step_label.html %} Crea y activa un entorno virtual dedicado para dbt.

  > **Importante:** En Windows se utiliza `C:\venvs\finops-dbt` para evitar problemas de instalación y `PATH` con Python instalado desde Microsoft Store.
  {: .lab-note .important .compact}

  ```powershell
  python -m venv C:\venvs\finops-dbt

  C:\venvs\finops-dbt\Scripts\Activate.ps1

  python -m pip install --upgrade pip

  python -m pip install dbt-postgres

  dbt --version
  ```

  > **Salida esperada:** El prompt muestra `(finops-dbt)` y `dbt --version` muestra dbt Core y el adapter PostgreSQL.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea `dbt\lab8\dbt_project.yml`.

  > **Nota:** El proyecto utiliza vistas para staging normalizado y una tabla para el mart final.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `dbt\lab8\dbt_project.yml`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```yaml
  name: finops_novatech
  version: "1.0"
  config-version: 2

  profile: finops_novatech

  model-paths:
    - models

  models:
    finops_novatech:
      staging:
        +materialized: view
      marts:
        +materialized: table
  ```

  > **Salida esperada:** El proyecto dbt queda configurado con modelos staging y marts.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea `dbt\lab8\profiles.yml`.

  > **Importante:** El schema base se establece en `public`; los modelos crearán objetos con nombres explícitos.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `dbt\lab8\profiles.yml`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```yaml
  finops_novatech:
    target: dev
    outputs:
      dev:
        type: postgres
        host: localhost
        port: 5432
        user: finops_admin
        password: finops2026
        dbname: finops_dw
        schema: public
        threads: 4
  ```

  > **Salida esperada:** dbt dispone de los datos necesarios para conectarse a PostgreSQL.
  {: .lab-note .output .compact}

- {% include step_label.html %} Valida la conexión de dbt con PostgreSQL.

  ```powershell
  Set-Location "$LabRoot\dbt\lab8"

  dbt debug --profiles-dir .
  ```

  > **Salida esperada:** dbt muestra `Connection test: [OK connection ok]` y `All checks passed!`.
  {: .lab-note .output .compact}

### Tarea 2.2. Crear modelos normalizados y mart unificado

- {% include step_label.html %} Crea los tres modelos staging normalizados.

  > **Nota:** El objetivo es mapear columnas nativas a un subconjunto común alineado con FOCUS 1.4.
  {: .lab-note .info .compact}

  > **Ubicación del código:** Crea cada archivo dentro de `dbt\lab8\models\staging`: `stg_aws_focus.sql`, `stg_azure_focus.sql` y `stg_gcp_focus.sql`. Pega en cada archivo únicamente el bloque SQL que corresponde a su nombre, desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```sql
  -- models/staging/stg_aws_focus.sql
  select
      line_item_usage_account_id::text as billing_account_id,
      'AWS Main Billing'::text as billing_account_name,
      date '2024-01-01' as billing_period_start,
      date '2024-01-31' as billing_period_end,
      left(line_item_usage_start_date, 10)::date as charge_period_start,
      left(line_item_usage_end_date, 10)::date as charge_period_end,
      case
          when line_item_line_item_type = 'Tax' then 'Tax'
          when line_item_line_item_type = 'Credit' then 'Credit'
          else 'Usage'
      end as charge_category,
      'Correction'::text as charge_class,
      'AWS'::text as service_provider_name,
      line_item_product_code::text as service_name,
      product_product_family::text as service_category,
      line_item_resource_id::text as resource_id,
      product_region_code::text as region_id,
      line_item_unblended_cost::numeric as billed_cost,
      line_item_net_unblended_cost::numeric as effective_cost,
      line_item_currency_code::text as billing_currency,
      nullif(resource_tags_user_team, '')::text as team,
      nullif(resource_tags_user_product, '')::text as product,
      nullif(resource_tags_user_environment, '')::text as environment,
      nullif(resource_tags_user_costcenter, '')::text as cost_center
  from staging.aws_costs
  ```

  ```sql
  -- models/staging/stg_azure_focus.sql
  select
      "BillingAccountId"::text as billing_account_id,
      'Azure Enterprise Billing'::text as billing_account_name,
      "BillingPeriodStartDate"::date as billing_period_start,
      "BillingPeriodEndDate"::date as billing_period_end,
      "Date"::date as charge_period_start,
      "Date"::date as charge_period_end,
      case
          when "ChargeType" = 'Tax' then 'Tax'
          when "ChargeType" = 'Credit' then 'Credit'
          when "ChargeType" = 'Purchase' then 'Purchase'
          else 'Usage'
      end as charge_category,
      'Correction'::text as charge_class,
      'Azure'::text as service_provider_name,
      "ServiceName"::text as service_name,
      "MeterCategory"::text as service_category,
      "ResourceId"::text as resource_id,
      "ResourceLocation"::text as region_id,
      "CostInBillingCurrency"::numeric as billed_cost,
      "EffectiveCost"::numeric as effective_cost,
      "BillingCurrency"::text as billing_currency,
      nullif("Tag_Team", '')::text as team,
      nullif("Tag_Product", '')::text as product,
      nullif("Tag_Environment", '')::text as environment,
      nullif("Tag_CostCenter", '')::text as cost_center
  from staging.azure_costs
  ```

  ```sql
  -- models/staging/stg_gcp_focus.sql
  select
      billing_account_id::text as billing_account_id,
      billing_account_name::text as billing_account_name,
      date '2024-01-01' as billing_period_start,
      date '2024-01-31' as billing_period_end,
      left(usage_start_time, 10)::date as charge_period_start,
      left(usage_end_time, 10)::date as charge_period_end,
      case
          when cost_type = 'tax' then 'Tax'
          when cost_type = 'adjustment' then 'Credit'
          else 'Usage'
      end as charge_category,
      'Correction'::text as charge_class,
      'GCP'::text as service_provider_name,
      "service.description"::text as service_name,
      "service.category"::text as service_category,
      "resource.name"::text as resource_id,
      "location.region"::text as region_id,
      cost::numeric as billed_cost,
      effective_cost::numeric as effective_cost,
      currency::text as billing_currency,
      nullif("labels.team", '')::text as team,
      nullif("labels.product", '')::text as product,
      nullif("labels.environment", '')::text as environment,
      nullif("labels.cost_center", '')::text as cost_center
  from staging.gcp_costs
  ```

  > **Salida esperada:** Existen tres modelos con nombres de columna homogéneos.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea `models\marts\focus_costs.sql` para unir y enriquecer los registros.

  > **Importante:** `team`, `product`, `environment` y `cost_center` son metadata organizacional derivada para análisis; no sustituyen la semántica completa de FOCUS Tags.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `dbt\lab8\models\marts\focus_costs.sql`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```sql
  select
      *,
      coalesce(team, 'Untagged') as allocation_team,
      coalesce(product, 'Untagged') as allocation_product,
      coalesce(environment, 'Untagged') as allocation_environment,
      coalesce(cost_center, 'Untagged') as allocation_cost_center,
      (
          case when team is not null then 1 else 0 end +
          case when product is not null then 1 else 0 end +
          case when environment is not null then 1 else 0 end +
          case when cost_center is not null then 1 else 0 end
      ) as metadata_fields_present,
      round(
          (
              (
                  case when team is not null then 1 else 0 end +
                  case when product is not null then 1 else 0 end +
                  case when environment is not null then 1 else 0 end +
                  case when cost_center is not null then 1 else 0 end
              )::numeric / 4
          ) * 100,
          1
      ) as metadata_completeness_pct
  from staging.focus_costs_base
  ```

  > **Salida esperada:** El mart unifica los tres proveedores y calcula metadata completeness de 0 a 100%.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea `scripts\build_focus_base.py` para consolidar los tres modelos staging.

  > **Importante:** Este paso evita depender de un `UNION ALL` dentro del modelo dbt final. Los tres modelos staging continúan siendo creados y normalizados por dbt.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\build_focus_base.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  import pandas as pd
  from sqlalchemy import create_engine

  engine = create_engine(
      "postgresql+psycopg2://finops_admin:finops2026@localhost:5432/finops_dw"
  )

  tables = [
      "stg_aws_focus",
      "stg_azure_focus",
      "stg_gcp_focus"
  ]

  frames = []

  for table in tables:
      df = pd.read_sql(
          f'SELECT * FROM public."{table}"',
          engine
      )
      frames.append(df)
      print(f"{table}: {len(df):,} registros")

  focus_base = pd.concat(
      frames,
      ignore_index=True
  )

  focus_base.to_sql(
      "focus_costs_base",
      engine,
      schema="staging",
      if_exists="replace",
      index=False
  )

  print(f"focus_costs_base: {len(focus_base):,} registros")
  ```

  > **Salida esperada:** Se crea `staging.focus_costs_base` con exactamente 3,300 registros.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta dbt para crear los modelos.

  > **Nota:** El mart final debe contener 3,300 registros.
  {: .lab-note .info .compact}

  ```powershell
  Set-Location "$LabRoot\dbt\lab8"

  dbt run --select stg_aws_focus stg_azure_focus stg_gcp_focus --profiles-dir .

  python -m pip install pandas
  python -m pip install sqlalchemy

  python "$LabRoot\scripts\build_focus_base.py"

  dbt run --select focus_costs --profiles-dir .
  ```

  > **Salida esperada:** Los tres modelos staging terminan correctamente, `focus_costs_base` contiene 3,300 registros y `focus_costs` se crea sin errores.
  {: .lab-note .output .compact}

{% capture r2 %}{{ results[1] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r2 %}

---

## 🧪 Tarea 3. Validar calidad y trazabilidad — 6 min

En esta tarea agregarás pruebas dbt mínimas para validar moneda, proveedor, fechas y campos críticos.

### Tarea 3.1. Crear pruebas de calidad

- {% include step_label.html %} Crea `dbt\lab8\models\schema.yml`.

  > **Importante:** Los tests validan integridad básica del modelo; no sustituyen controles completos de Data Quality.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `dbt\lab8\models\schema.yml`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```yaml
  version: 2

  models:
    - name: focus_costs
      columns:
        - name: billing_account_id
          tests:
            - not_null

        - name: charge_period_start
          tests:
            - not_null

        - name: service_provider_name
          tests:
            - not_null
            - accepted_values:
                values: ['AWS', 'Azure', 'GCP']

        - name: billing_currency
          tests:
            - not_null
            - accepted_values:
                values: ['EUR']

        - name: billed_cost
          tests:
            - not_null

        - name: allocation_team
          tests:
            - accepted_values:
                values:
                  - Backend
                  - Frontend
                  - Data
                  - Platform
                  - Untagged
  ```

  > **Salida esperada:** El archivo define validaciones sobre columnas críticas.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta los tests y consulta el total por proveedor.

  > **Nota:** Todos los tests deben completar sin errores antes de visualizar los datos.
  {: .lab-note .info .compact}

  ```powershell
  dbt test --profiles-dir .

  docker exec finops-postgres psql `
    -U finops_admin `
    -d finops_dw `
    -c "SELECT service_provider_name, COUNT(*) AS records, ROUND(SUM(effective_cost)::numeric, 2) AS effective_cost_eur FROM public.focus_costs GROUP BY service_provider_name ORDER BY effective_cost_eur DESC;"
  ```

  > **Salida esperada:** Los tests pasan y la consulta muestra AWS, Azure y GCP con sus totales.
  {: .lab-note .output .compact}

### Tarea 3.2. Validar metadata completeness

- {% include step_label.html %} Consulta la calidad de metadata por proveedor.

  > **Nota:** El porcentaje representa completitud de Team, Product, Environment y CostCenter por registro de costo.
  {: .lab-note .info .compact}

  ```powershell
  docker exec finops-postgres psql `
    -U finops_admin `
    -d finops_dw `
    -c "SELECT service_provider_name, ROUND(AVG(metadata_completeness_pct),1) AS avg_metadata_pct FROM public.focus_costs GROUP BY service_provider_name ORDER BY service_provider_name;"
  ```

  > **Salida esperada:** Cada proveedor muestra un porcentaje promedio de metadata completeness.
  {: .lab-note .output .compact}

{% capture r3 %}{{ results[2] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r3 %}

---

## 📊 Tarea 4. Visualizar y documentar — 7 min

En esta tarea conectarás Grafana al mart normalizado, crearás un dashboard con cuatro paneles y documentarás el gobierno básico del pipeline.

### Tarea 4.1. Configurar Grafana y dashboard

- {% include step_label.html %} Crea `outputs\lab8\grafana_dashboard.json`.

  > **Nota:** El dashboard utiliza cuatro vistas: servicio, equipo, tendencia y metadata completeness.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `outputs\lab8\grafana_dashboard.json`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```json
  {
    "dashboard": {
      "title": "NovaTech FinOps - FOCUS 1.4 Aligned",
      "tags": ["finops", "focus", "novatech"],
      "timezone": "browser",
      "time": {
        "from": "2024-01-01T00:00:00.000Z",
        "to": "2024-02-01T00:00:00.000Z"
      },
      "panels": [
        {
          "id": 1,
          "title": "EffectiveCost por Servicio",
          "type": "barchart",
          "gridPos": {"h": 9, "w": 12, "x": 0, "y": 0},
          "targets": [{
            "rawSql": "SELECT service_name, ROUND(SUM(effective_cost)::numeric, 2) AS effective_cost_eur FROM public.focus_costs GROUP BY service_name ORDER BY effective_cost_eur DESC LIMIT 10",
            "format": "table"
          }]
        },
        {
          "id": 2,
          "title": "EffectiveCost por Equipo",
          "type": "piechart",
          "gridPos": {"h": 9, "w": 12, "x": 12, "y": 0},
          "targets": [{
            "rawSql": "SELECT allocation_team, ROUND(SUM(effective_cost)::numeric, 2) AS effective_cost_eur FROM public.focus_costs GROUP BY allocation_team ORDER BY effective_cost_eur DESC",
            "format": "table"
          }]
        },
        {
          "id": 3,
          "title": "Tendencia Diaria por Proveedor",
          "type": "timeseries",
          "gridPos": {"h": 9, "w": 16, "x": 0, "y": 9},
          "targets": [{
            "rawSql": "SELECT charge_period_start AS time, service_provider_name AS metric, SUM(effective_cost) AS value FROM public.focus_costs GROUP BY charge_period_start, service_provider_name ORDER BY time",
            "format": "time_series"
          }]
        },
        {
          "id": 4,
          "title": "% Registros con Metadata Incompleta",
          "type": "stat",
          "gridPos": {"h": 9, "w": 8, "x": 16, "y": 9},
          "targets": [{
            "rawSql": "SELECT ROUND((COUNT(*) FILTER (WHERE metadata_completeness_pct < 100)::numeric / COUNT(*)::numeric) * 100,1) AS pct_incomplete FROM public.focus_costs",
            "format": "table"
          }]
        }
      ],
      "schemaVersion": 39
    },
    "overwrite": true
  }
  ```

  > **Salida esperada:** Se crea el archivo JSON del dashboard.
  {: .lab-note .output .compact}

- {% include step_label.html %} Configura el datasource y carga el dashboard mediante la API de Grafana.

  > **Importante:** Grafana está disponible en `http://localhost:3000` con las credenciales definidas en Docker Compose. `admin`, `finops2026`. Cambiar manualmente **Tendencia Diaria por Proveedor** a Enero 2024 en el panel de Grafana.
  {: .lab-note .important .compact}

  ```powershell
  $Credential = [Convert]::ToBase64String(
    [Text.Encoding]::ASCII.GetBytes("admin:finops2026")
  )

  $Headers = @{
    Authorization = "Basic $Credential"
    "Content-Type" = "application/json"
  }

  $DatasourceBody = @{
    name = "FinOps PostgreSQL"
    type = "postgres"
    url = "postgres:5432"
    database = "finops_dw"
    user = "finops_admin"
    access = "proxy"
    isDefault = $true
    secureJsonData = @{
      password = "finops2026"
    }
    jsonData = @{
      sslmode = "disable"
      postgresVersion = 1600
    }
  } | ConvertTo-Json -Depth 5

  Invoke-RestMethod `
    -Method Post `
    -Uri "http://localhost:3000/api/datasources" `
    -Headers $Headers `
    -Body $DatasourceBody

  $DashboardBody = Get-Content `
    "$LabRoot\outputs\lab8\grafana_dashboard.json" `
    -Raw

  Invoke-RestMethod `
    -Method Post `
    -Uri "http://localhost:3000/api/dashboards/db" `
    -Headers $Headers `
    -Body $DashboardBody
  ```

  > **Salida esperada:** Grafana crea el datasource y registra el dashboard NovaTech.
  {: .lab-note .output .compact}

### Tarea 4.2. Documentar gobierno de datos y validar

- {% include step_label.html %} Crea `docs\lab8\finops_data_governance.md`.

  > **Nota:** El documento resume ownership, linaje, freshness y reglas de calidad.
  {: .lab-note .info .compact}

  > **Ubicación del contenido:** En `docs\lab8\finops_data_governance.md`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```markdown
  # FinOps Data Governance - NovaTech SRL

  ## Ownership
  - Raw ingestion: Platform
  - FOCUS-aligned normalization: FinOps
  - Allocation metadata: FinOps + Product/Engineering
  - Visualization: FinOps

  ## Lineage
  AWS / Azure / GCP raw CSV
  → PostgreSQL staging
  → dbt normalized models
  → focus_costs mart
  → Grafana

  ## NovaTech Freshness Target
  - Maximum data freshness target: 48 hours
  - Pipeline frequency: daily

  ## Data Quality Rules
  - billing_account_id must not be null
  - billing_currency must be EUR
  - service_provider_name must be AWS, Azure or GCP
  - billed_cost must not be null
  - allocation_team may use Untagged when metadata is missing
  ```

  > **Salida esperada:** Se crea la documentación de gobierno del pipeline.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta una validación final del laboratorio.

  > **Importante:** El mart debe contener exactamente 3,300 registros.
  {: .lab-note .important .compact}

  ```powershell
  docker exec finops-postgres psql `
    -U finops_admin `
    -d finops_dw `
    -c "SELECT COUNT(*) AS total_records FROM public.focus_costs;"

  Get-Item `
    "$LabRoot\outputs\lab8\grafana_dashboard.json", `
    "$LabRoot\docs\lab8\finops_data_governance.md" |
    Select-Object Name, Length

  docker compose `
    -f "$LabRoot\docker\lab8\docker-compose.yml" `
    ps
  ```

  > **Salida esperada:** `focus_costs` contiene 3,300 registros, los dos artefactos existen y PostgreSQL/Grafana continúan activos.
  {: .lab-note .output .compact}

{% capture r4 %}{{ results[3] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r4 %}
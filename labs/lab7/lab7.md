---
layout: lab
title: "Demostración 7: Definición de roles, RACI y políticas mínimas"
permalink: /lab7/lab7/
images_base: /labs/lab7/img
duration: "25 minutos"
objective:
  - Construir los artefactos mínimos de Governance para NovaTech SRL mediante definición de personas y responsabilidades, matriz RACI, políticas operativas y un Governance Readiness Assessment.
prerequisites:
  - Equipo con Windows 11 y Visual Studio Code instalado.
  - Python 3.12 o superior disponible desde la terminal integrada de Visual Studio Code.
  - Paquetes pandas, openpyxl y pyyaml instalados.
  - Haber completado la Práctica 6 o disponer de `finops_optimization_backlog.xlsx`.
  - Conocimiento básico de RACI, Governance, Policy & Risk y responsabilidades FinOps.
introduction:
  - En esta demostración definirás cómo NovaTech organiza la toma de decisiones FinOps. Mapearás personas y equipos, construirás una matriz RACI, generarás cuatro políticas YAML y evaluarás la preparación de Governance mediante cinco dimensiones operativas.
slug: lab7
lab_number: 7
final_result: >
  Al finalizar tendrás un archivo finops_governance_v1.xlsx con cuatro hojas, cuatro políticas YAML y un assessment de Governance que permitirá identificar gaps de accountability, policy, enforcement, exception management y review & risk.
notes:
  - Todos los archivos se almacenan dentro de la carpeta finops-course del Escritorio de Windows.
  - Todos los importes monetarios se expresan en EUR.
  - Las personas y responsabilidades utilizadas representan la adaptación organizacional de NovaTech y no una lista cerrada de roles oficiales de FinOps.
  - Los thresholds, SLA y reglas de aprobación son políticas internas de NovaTech.
  - El Governance Readiness Assessment es un instrumento interno y no debe interpretarse como un score oficial de madurez Crawl, Walk o Run.
references:
  - text: "FinOps Personas"
    url: https://www.finops.org/framework/personas/
  - text: "FinOps Governance, Policy & Risk"
    url: https://www.finops.org/framework/capabilities/governance-policy-risk/
  - text: "FinOps Maturity Model"
    url: https://www.finops.org/framework/maturity-model/
  - text: "PyYAML Documentation"
    url: https://pyyaml.org/wiki/PyYAMLDocumentation
prev: /lab6/lab6/
next: /lab8/lab8/
---

---

<!-- Aquí comienzan las instrucciones paso a paso de la práctica -->

## 👥 Tarea 1. Preparar Governance y definir responsabilidades — 6 min

En esta tarea descargarás el perfil de Governance de NovaTech, validarás su estructura y definirás las personas y responsabilidades que participarán en los procesos FinOps.

### Tarea 1.1. Preparar el entorno y descargar el perfil

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
    "$LabRoot\outputs\lab7", `
    "$LabRoot\policies\lab7", `
    "$LabRoot\scripts" | Out-Null

  Set-Location $LabRoot
  ```

  > **Salida esperada:** Existen `data\raw`, `data\processed`, `outputs\lab7`, `policies\lab7` y `scripts`.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define la URL del perfil de Governance y descárgalo.

  ```powershell
  $ProfileUrl = "https://d3gfd4cq9u8k8p.cloudfront.net/courses/finops/novatech_governance_profile.json"

  Invoke-WebRequest `
    -Uri $ProfileUrl `
    -OutFile "$LabRoot\data\raw\novatech_governance_profile.json"
  ```

  > **Salida esperada:** Se crea `data\raw\novatech_governance_profile.json`.
  {: .lab-note .output .compact}

### Tarea 1.2. Cargar perfil y definir personas FinOps

- {% include step_label.html %} Verifica que PyYAML esté instalado en el mismo Python que ejecutará el script.

  > **Importante:** El módulo se importa en Python como `yaml`, pero el paquete que debe instalarse con pip se llama `PyYAML`. No uses `pip install yaml`.
  {: .lab-note .important .compact}

  ```powershell
  python -m pip install PyYAML
  python -c "import yaml; print('PyYAML:', yaml.__version__)"
  ```

  > **Salida esperada:** PowerShell confirma la instalación y muestra la versión de PyYAML sin errores.
  {: .lab-note .output .compact}

- {% include step_label.html %} Crea `scripts\governance_analysis.py` y carga el perfil organizacional.

  > **Nota:** El perfil contiene equipos, responsables, políticas base y objetivos de Governance de NovaTech.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, pega este bloque desde la **primera línea del archivo**.
  {: .lab-note .important .compact}

  ```python
  from pathlib import Path
  import json
  import pandas as pd
  import yaml

  ROOT = Path.home() / "Desktop" / "finops-course"
  RAW = ROOT / "data" / "raw"
  PROCESSED = ROOT / "data" / "processed"
  OUTPUTS = ROOT / "outputs" / "lab7"
  POLICIES = ROOT / "policies" / "lab7"

  OUTPUTS.mkdir(parents=True, exist_ok=True)
  POLICIES.mkdir(parents=True, exist_ok=True)

  with open(
      RAW / "novatech_governance_profile.json",
      "r",
      encoding="utf-8"
  ) as f:
      profile = json.load(f)

  print(f"Empresa: {profile['company']}")
  print(f"Equipos: {len(profile['teams'])}")
  print(f"Productos: {len(profile['products'])}")
  print(f"Moneda: {profile['currency']}")
  ```

  > **Salida esperada:** Se muestran NovaTech SRL, 8 equipos, 3 productos y EUR.
  {: .lab-note .output .compact}

- {% include step_label.html %} Define las seis personas organizacionales utilizadas por NovaTech.

  > **Importante:** Estas personas representan la adaptación de NovaTech para organizar responsabilidades FinOps.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo de** `print(f"Moneda: {profile['currency']}")`.
  {: .lab-note .important .compact}

  ```python
  roles_data = [
      {
          "Persona": "Finanzas",
          "Teams": "Finance",
          "Lead": profile["teams"]["Finance"]["lead"],
          "PrimaryResponsibility": "Budget, reporting, reconciliation and allocation oversight",
          "Cadence": "Weekly"
      },
      {
          "Persona": "Ingeniería",
          "Teams": "Backend, Frontend, Data",
          "Lead": profile["teams"]["Backend"]["lead"],
          "PrimaryResponsibility": "Usage optimization, metadata and remediation execution",
          "Cadence": "Daily"
      },
      {
          "Persona": "Producto",
          "Teams": "Product",
          "Lead": profile["teams"]["Product"]["lead"],
          "PrimaryResponsibility": "Business value, prioritization and unit economics context",
          "Cadence": "Biweekly"
      },
      {
          "Persona": "Operaciones",
          "Teams": "Platform",
          "Lead": profile["teams"]["Platform"]["lead"],
          "PrimaryResponsibility": "Alerts, automation, inventory and platform controls",
          "Cadence": "Daily"
      },
      {
          "Persona": "Compras",
          "Teams": "Procurement",
          "Lead": profile["teams"]["Procurement"]["lead"],
          "PrimaryResponsibility": "Commercial commitments, contracts and rate optimization support",
          "Cadence": "Monthly"
      },
      {
          "Persona": "Liderazgo",
          "Teams": "Leadership",
          "Lead": profile["teams"]["Leadership"]["lead"],
          "PrimaryResponsibility": "Sponsorship, escalation and strategic decision making",
          "Cadence": "Monthly"
      }
  ]

  roles_df = pd.DataFrame(roles_data)

  print(
      roles_df[
          ["Persona", "Teams", "Lead"]
      ].to_string(index=False)
  )
  ```

  > **Salida esperada:** Se muestran seis personas organizacionales con equipos y responsables asignados.
  {: .lab-note .output .compact}

{% assign results = site.data.task-results[page.slug].results %}
{% capture r1 %}{{ results[0] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r1 %}

---

## 🧭 Tarea 2. Construir la matriz RACI — 6 min

En esta tarea asignarás accountability y responsabilidad a cinco procesos FinOps clave y validarás que cada proceso tenga exactamente un Accountable.

### Tarea 2.1. Definir la matriz RACI

- {% include step_label.html %} Agrega la matriz RACI de NovaTech al script.

  > **Nota:** `R` ejecuta, `A` responde por el resultado, `C` es consultado e `I` es informado.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo del bloque que imprime** `roles_df`.
  {: .lab-note .important .compact}

  ```python
  raci_data = {
      "Process": [
          "Cost Allocation",
          "Monthly Cost Reporting",
          "Commitment Approval",
          "Cost Anomaly Management",
          "Optimization Backlog Review"
      ],
      "Finanzas": ["A", "R", "C", "I", "C"],
      "Ingeniería": ["R", "C", "C", "R", "R"],
      "Producto": ["C", "I", "I", "I", "A"],
      "Operaciones": ["R", "R", "I", "R", "C"],
      "Compras": ["I", "I", "R", "I", "I"],
      "Liderazgo": ["I", "A", "A", "I", "I"]
  }

  raci_df = pd.DataFrame(raci_data).set_index("Process")

  print(raci_df.to_string())
  ```

  > **Salida esperada:** Se muestran cinco procesos FinOps y seis columnas de personas.
  {: .lab-note .output .compact}

- {% include step_label.html %} Valida que cada proceso tenga un `A` y al menos un `R`.

  > **Importante:** Más de un Accountable puede generar ambigüedad de decisión; esta práctica exige uno por proceso.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo de** `print(raci_df.to_string())`.
  {: .lab-note .important .compact}

  ```python
  validation_rows = []

  for process in raci_df.index:
      row = raci_df.loc[process]
      accountable_count = int((row == "A").sum())
      responsible_count = int((row == "R").sum())

      validation_rows.append({
          "Process": process,
          "AccountableCount": accountable_count,
          "ResponsibleCount": responsible_count,
          "Valid": (
              accountable_count == 1
              and responsible_count >= 1
          )
      })

  raci_validation = pd.DataFrame(validation_rows)

  print(raci_validation.to_string(index=False))
  ```

  > **Salida esperada:** Todos los procesos muestran `AccountableCount = 1`, al menos un Responsible y `Valid = True`.
  {: .lab-note .output .compact}

### Tarea 2.2. Relacionar backlog y responsabilidad

- {% include step_label.html %} Carga el backlog de optimización de la Práctica 6.

  > **Nota:** El backlog permite conectar Governance con acciones FinOps reales.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo de** `print(raci_validation.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  backlog_path = PROCESSED / "finops_optimization_backlog.xlsx"

  backlog_df = pd.read_excel(
      backlog_path,
      sheet_name="backlog"
  )

  print(f"Iniciativas cargadas: {len(backlog_df)}")
  ```

  > **Salida esperada:** Se cargan las oportunidades generadas en la Práctica 6.
  {: .lab-note .output .compact}

- {% include step_label.html %} Asigna una persona responsable según `ActionType`.

  > **Nota:** Este mapping simplifica la responsabilidad principal de cada tipo de acción.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo de** `print(f"Iniciativas cargadas: {len(backlog_df)}")`.
  {: .lab-note .important .compact}

  ```python
  action_owner = {
      "Efficiency": "Ingeniería",
      "Elimination": "Operaciones",
      "Rate": "Compras",
      "Governance": "Operaciones"
  }

  backlog_df["ResponsiblePersona"] = (
      backlog_df["ActionType"]
      .map(action_owner)
      .fillna("Finanzas")
  )

  print(
      backlog_df[
          [
              "BacklogId",
              "OptimizationCategory",
              "ActionType",
              "ResponsiblePersona"
          ]
      ]
      .head(10)
      .to_string(index=False)
  )
  ```

  > **Salida esperada:** Las primeras iniciativas muestran una persona responsable asignada.
  {: .lab-note .output .compact}

{% capture r2 %}{{ results[1] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r2 %}

---

## 📜 Tarea 3. Definir políticas operativas — 7 min

En esta tarea crearás cuatro políticas YAML para metadata, presupuesto, excepciones y aprobaciones de gasto.

### Tarea 3.1. Crear políticas de metadata y presupuesto

- {% include step_label.html %} Agrega la política de metadata obligatoria.

  > **Importante:** Las dimensiones coinciden con las utilizadas en la Práctica 4 para mantener consistencia de Allocation.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo del bloque que imprime las primeras iniciativas con** `ResponsiblePersona`.
  {: .lab-note .important .compact}

  ```python
  policy_tagging = {
      "policy_id": "POL-TAG-001",
      "policy_name": "Mandatory Allocation Metadata",
      "version": "1.0",
      "currency": "EUR",
      "required_metadata": [
          "Environment",
          "Team",
          "Product",
          "CostCenter"
      ],
      "optional_metadata": [
          "Owner"
      ],
      "compliance_target_pct": (
          profile["governance_targets"]
          ["mandatory_metadata_compliance_pct"]
      ),
      "non_compliance_action": (
          "Notify owner and create remediation item"
      )
  }
  ```

  > **Salida esperada:** Se define una política con cuatro dimensiones obligatorias y `Owner` opcional.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega la política de Budget y alertas.

  > **Nota:** Los thresholds RAG provienen de la política interna utilizada en la Práctica 5.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo de** la definición de `policy_tagging`.
  {: .lab-note .important .compact}

  ```python
  policy_budget = {
      "policy_id": "POL-BUD-001",
      "policy_name": "Budget Monitoring",
      "version": "1.0",
      "currency": "EUR",
      "monthly_budget_total_eur": (
          profile["budget_policy"]
          ["monthly_budget_total_eur"]
      ),
      "thresholds": (
          profile["budget_policy"]["thresholds"]
      ),
      "review_cadence": (
          profile["budget_policy"]["review_cadence"]
      ),
      "escalation": {
          "amber": "Team Lead + Finance",
          "red": "Leadership + Finance"
      }
  }
  ```

  > **Salida esperada:** Se define presupuesto mensual, thresholds y reglas de escalamiento.
  {: .lab-note .output .compact}

### Tarea 3.2. Crear excepciones y aprobaciones

- {% include step_label.html %} Agrega la política de excepciones.

  > **Nota:** Una excepción debe ser temporal, justificada y revisable.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo de** la definición de `policy_budget`.
  {: .lab-note .important .compact}

  ```python
  policy_exceptions = {
      "policy_id": "POL-EXC-001",
      "policy_name": "FinOps Exception Management",
      "version": "1.0",
      "approval_sla_hours": (
          profile["exception_policy"]
          ["approval_sla_hours"]
      ),
      "max_duration_days": (
          profile["exception_policy"]
          ["max_duration_days"]
      ),
      "renewal_required": (
          profile["exception_policy"]
          ["renewal_required"]
      ),
      "required_fields": [
          "BusinessJustification",
          "EstimatedImpactEUR",
          "ExceptionDurationDays",
          "RemediationPlan",
          "Owner"
      ]
  }
  ```

  > **Salida esperada:** La política exige justificación, impacto, duración, plan de remediación y owner.
  {: .lab-note .output .compact}

- {% include step_label.html %} Agrega la política de aprobaciones de gasto.

  > **Importante:** Los importes y SLA son thresholds internos de NovaTech.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo de** la definición de `policy_exceptions`.
  {: .lab-note .important .compact}

  ```python
  policy_approvals = {
      "policy_id": "POL-APR-001",
      "policy_name": "Technology Spend Approvals",
      "version": "1.0",
      "currency": "EUR",
      "thresholds": (
          profile["approval_policy"]
          ["thresholds_eur_monthly"]
      ),
      "commitment_rule": (
          "Commitments require Finance review before approval"
      )
  }
  ```

  > **Salida esperada:** La política contiene cuatro rangos de aprobación expresados en EUR.
  {: .lab-note .output .compact}

- {% include step_label.html %} Guarda las cuatro políticas como archivos YAML.

  > **Nota:** YAML permite mantener políticas legibles y versionables junto con otros artefactos técnicos.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo de** la definición de `policy_approvals`.
  {: .lab-note .important .compact}

  ```python
  policies = {
      "policy_tagging.yaml": policy_tagging,
      "policy_budgets.yaml": policy_budget,
      "policy_exceptions.yaml": policy_exceptions,
      "policy_approvals.yaml": policy_approvals
  }

  for filename, policy in policies.items():
      filepath = POLICIES / filename

      with open(filepath, "w", encoding="utf-8") as f:
          yaml.safe_dump(
              policy,
              f,
              allow_unicode=True,
              sort_keys=False
          )

      print(f"Creado: {filepath}")
  ```

  > **Salida esperada:** Se crean cuatro archivos `.yaml` en `policies\lab7`.
  {: .lab-note .output .compact}

{% capture r3 %}{{ results[2] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r3 %}

---

## ✅ Tarea 4. Evaluar Governance Readiness y exportar — 6 min

En esta tarea evaluarás cinco dimensiones de Governance y consolidarás Roles, RACI, Policies y Assessment en un archivo Excel.

### Tarea 4.1. Calcular Governance Readiness

- {% include step_label.html %} Agrega las cinco dimensiones del assessment.

  > **Importante:** Este assessment es interno de NovaTech y no representa una clasificación oficial Crawl, Walk o Run.
  {: .lab-note .important .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo del bloque que crea los cuatro archivos YAML**.
  {: .lab-note .important .compact}

  ```python
  governance_assessment = {
      "Accountability": [
          ("RACI documented", True),
          ("Single accountable per process", True),
          ("Owners assigned to optimization backlog", True),
          ("Escalation owner defined", False)
      ],
      "Policy": [
          ("Metadata policy documented", True),
          ("Budget policy documented", True),
          ("Exception policy documented", True),
          ("Spend approval policy documented", True)
      ],
      "Enforcement": [
          ("Metadata compliance checked automatically", False),
          ("Budget alerts automated", False),
          ("Approval workflow automated", False),
          ("Policy violations create remediation items", False)
      ],
      "Exception Management": [
          ("Exception SLA defined", True),
          ("Maximum exception duration defined", True),
          ("Renewal required", True),
          ("Exception register maintained", False)
      ],
      "Review & Risk": [
          ("Weekly budget review defined", True),
          ("Quarterly policy review defined", True),
          ("Monthly exception review defined", True),
          ("Risk review integrated with leadership", False)
      ]
  }

  assessment_rows = []

  for dimension, checks in governance_assessment.items():
      yes_count = sum(1 for _, answer in checks if answer)
      total = len(checks)

      assessment_rows.append({
          "Dimension": dimension,
          "Controls": total,
          "Implemented": yes_count,
          "ReadinessPct": round(
              yes_count / total * 100,
              1
          )
      })

  assessment_df = pd.DataFrame(assessment_rows)

  print(assessment_df.to_string(index=False))
  ```

  > **Salida esperada:** Se muestran cinco dimensiones con porcentaje de readiness.
  {: .lab-note .output .compact}

- {% include step_label.html %} Identifica las dimensiones con menor readiness.

  > **Nota:** El objetivo es detectar gaps prioritarios, no asignar una etiqueta global de madurez.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo de** `print(assessment_df.to_string(index=False))`.
  {: .lab-note .important .compact}

  ```python
  gaps = assessment_df.sort_values(
      "ReadinessPct"
  ).reset_index(drop=True)

  print("Prioridad de mejora:")
  print(
      gaps[
          ["Dimension", "ReadinessPct"]
      ].to_string(index=False)
  )
  ```

  > **Salida esperada:** Las dimensiones se muestran ordenadas desde menor a mayor readiness.
  {: .lab-note .output .compact}

### Tarea 4.2. Exportar y validar Governance

- {% include step_label.html %} Crea el resumen de políticas y exporta `finops_governance_v1.xlsx`.

  > **Nota:** El Excel consolida las vistas principales; los YAML permanecen como artefactos versionables separados.
  {: .lab-note .info .compact}

  > **Ubicación del código:** En `scripts\governance_analysis.py`, agrega este bloque **debajo del bloque que imprime** `Prioridad de mejora`.
  {: .lab-note .important .compact}

  ```python
  policies_summary = pd.DataFrame([
      {
          "PolicyId": policy_tagging["policy_id"],
          "PolicyName": policy_tagging["policy_name"],
          "Owner": "Platform",
          "Review": "Monthly"
      },
      {
          "PolicyId": policy_budget["policy_id"],
          "PolicyName": policy_budget["policy_name"],
          "Owner": "Finance",
          "Review": profile["budget_policy"]["review_cadence"]
      },
      {
          "PolicyId": policy_exceptions["policy_id"],
          "PolicyName": policy_exceptions["policy_name"],
          "Owner": "Leadership",
          "Review": profile["governance_targets"]["exception_review_frequency"]
      },
      {
          "PolicyId": policy_approvals["policy_id"],
          "PolicyName": policy_approvals["policy_name"],
          "Owner": "Leadership",
          "Review": profile["governance_targets"]["policy_review_frequency"]
      }
  ])

  report_xlsx = PROCESSED / "finops_governance_v1.xlsx"

  with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
      roles_df.to_excel(writer, sheet_name="roles", index=False)
      raci_df.reset_index().to_excel(writer, sheet_name="raci", index=False)
      policies_summary.to_excel(writer, sheet_name="policies", index=False)
      assessment_df.to_excel(
          writer,
          sheet_name="governance_assessment",
          index=False
      )

  print(f"Reporte creado: {report_xlsx}")
  ```

  > **Salida esperada:** Se crea `data\processed\finops_governance_v1.xlsx` con cuatro hojas.
  {: .lab-note .output .compact}

- {% include step_label.html %} Ejecuta el script completo y valida los entregables.

  > **Importante:** El Excel y los cuatro YAML deben existir y poder leerse correctamente.
  {: .lab-note .important .compact}

  ```powershell
  python "$LabRoot\scripts\governance_analysis.py"
  ```
  ```powershell
  Get-Item `
      "$LabRoot\data\processed\finops_governance_v1.xlsx", `
      "$LabRoot\policies\lab7\policy_tagging.yaml", `
      "$LabRoot\policies\lab7\policy_budgets.yaml", `
      "$LabRoot\policies\lab7\policy_exceptions.yaml", `
      "$LabRoot\policies\lab7\policy_approvals.yaml" |
      Select-Object Name, Length
  ```
  ```powershell
  python -c "import pandas as pd; p=r'$LabRoot\data\processed\finops_governance_v1.xlsx'; x=pd.ExcelFile(p); print(x.sheet_names)"
  ```

  > **Salida esperada:** Los cinco archivos tienen tamaño mayor a 0 bytes y el Excel contiene `roles`, `raci`, `policies` y `governance_assessment`.
  {: .lab-note .output .compact}

{% capture r4 %}{{ results[3] }}{% endcapture %}
{% include task-result.html title="Tarea finalizada" content=r4 %}
# ReporteSAC — Sistema de Reportes del Seguro Agrícola Catastrófico

![.NET](https://img.shields.io/badge/.NET-8.0-512BD4?logo=dotnet)
![Python](https://img.shields.io/badge/Python-3.11-3776AB?logo=python&logoColor=white)
![Docker](https://img.shields.io/badge/Docker-compose-2496ED?logo=docker&logoColor=white)
![License](https://img.shields.io/github/license/ivanmena2021/app-sac-net)
![Last commit](https://img.shields.io/github/last-commit/ivanmena2021/app-sac-net)
![Repo size](https://img.shields.io/github/repo-size/ivanmena2021/app-sac-net)

Sistema web para el monitoreo y generación de reportes del **Seguro Agrícola Catastrófico (SAC) 2025-2026**, desarrollado para la Dirección de Seguro y Fomento de Financiamiento Agrario (DSFFA) — **MIDAGRI**.

---

## ✨ Características destacadas

- 🤖 **Consultas en lenguaje natural** sobre los datos SAC usando Claude API (módulo `LlmQuery`)
- 🔄 **Descarga automática** desde portales de aseguradoras (Rimac SISGAQSAC, Agroevaluaciones)
- 📄 **Generación multi-formato** — Word, Excel, PDF y PowerPoint desde un solo origen de datos
- 🗂️ **Sesiones aisladas por usuario** (`UploadedFileStore` scoped) — varios usuarios trabajan en paralelo sin mezclar datos
- 🚦 **Sistema de semáforos** para alertas departamentales
- 🧩 **Clean Architecture** — Domain / Application / Infrastructure / WebApp

---

## 🏗️ Arquitectura

El sistema consta de dos servicios que trabajan juntos:

```
Navegador (usuario)
    |
    v
WebApp (.NET 8 Blazor Server)    -- UI, carga de datos, lógica de negocio
    |
    v  HTTP (multipart/form-data)
Python API (FastAPI)              -- Generación de reportes Word, Excel, PDF, PPT
```

### WebApp (.NET 8)
- **Framework:** ASP.NET Core 8, Blazor Server (InteractiveServer)
- **Arquitectura:** Clean Architecture (Domain, Application, Infrastructure, WebApp)
- **Páginas:** Dashboard, Reportes, Departamental, Semáforo, Mapa de Calor, Calendario, Clima y Riesgo, Comparativo, Explorar, Consultas LLM

### Python API (FastAPI)
- **Framework:** FastAPI + Uvicorn
- **Generadores:** Word (python-docx), Excel (openpyxl), PDF (fpdf2), PowerPoint (python-pptx)
- **Datos estáticos:** Consolidados históricos, materia asegurada, primas, perfiles de riesgo

## 📁 Estructura del proyecto

```
ReporteSAC/
├── ReporteSAC.sln
├── Dockerfile                    # Build .NET app
├── docker-compose.yml            # Orquesta ambos servicios
├── docs/
│   └── GUIA_DESPLIEGUE_MIDAGRI.md
├── src/
│   ├── Domain/                   # Entidades del dominio (SiniestroRecord, etc.)
│   ├── Application/              # DTOs, interfaces de servicios, lógica de procesamiento
│   │   ├── Contracts/Services/   # IWordReportService, IExcelReportService, etc.
│   │   ├── DTOs/                 # DatosNacionalesDto, DatosDepartamentalDto, etc.
│   │   └── Services/             # DataProcessorService
│   ├── Infrastructure/           # Implementaciones
│   │   ├── PythonApi/            # PythonApiReportService (HttpClient → Python API)
│   │   ├── Session/              # UploadedFileStore (scoped per-user)
│   │   ├── Alertas/              # SemaforoService
│   │   ├── AutoDownload/         # Descarga automática desde portales aseguradores
│   │   ├── LlmQuery/             # Consultas en lenguaje natural (Claude API)
│   │   └── ExcelReader/          # Lectura de archivos Excel
│   └── WebApp/                   # Blazor Server app
│       ├── Components/Pages/     # 12 páginas Razor
│       ├── wwwroot/              # CSS, JS, iconos
│       └── Program.cs            # Entry point
└── python-api/
    ├── Dockerfile
    ├── main.py                   # FastAPI app con endpoints de generación
    ├── requirements.txt
    ├── data_processor.py         # Procesamiento de datos para generadores
    ├── gen_word_nacional_py.py   # Ayuda memoria nacional
    ├── gen_word_departamental_py.py  # Ayuda memoria departamental
    ├── gen_word_operatividad.py  # Reporte de operatividad
    ├── gen_excel_eme.py          # Reporte EME
    ├── gen_excel_enhanced.py     # Excel mejorado con gráficos
    ├── gen_pdf_resumen.py        # Resumen ejecutivo PDF
    ├── gen_ppt_dinamico.py       # PPT dinámica con filtros
    ├── gen_ppt_historico.py      # PPT histórica por departamento
    └── static_data/              # Archivos estáticos requeridos
```

## 📊 Reportes generados

| Reporte | Formato | Descripción |
|---------|---------|-------------|
| Ayuda Memoria Nacional | .docx | Resumen ejecutivo nacional con tablas y métricas |
| Ayuda Memoria Departamental | .docx | Detalle por departamento con provincias y eventos |
| Operatividad | .docx | Estado operativo del SAC por departamento |
| Reporte EME | .xlsx | Formato oficial para la Oficina de Emergencias |
| Excel Mejorado | .xlsx | Consolidado con gráficos y múltiples hojas |
| Resumen Ejecutivo | .pdf | Resumen visual con gráficos y KPIs |
| PPT Dinámica | .pptx | Presentación con filtros por empresa, zona, período |
| PPT Histórica | .pptx | Comparativo histórico por departamento |

## 🚀 Inicio rápido

### Opción 1: Docker Compose (recomendado)

```bash
git clone https://github.com/ivanmena2021/app-sac-net.git
cd app-sac-net
docker compose up --build -d
```

Acceder a: http://localhost:5141

### Opción 2: Desarrollo local

**Terminal 1 — Python API:**
```bash
cd python-api
python -m pip install -r requirements.txt
python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

**Terminal 2 — WebApp .NET:**
```bash
dotnet run --project src/WebApp
```

Acceder a: https://localhost:5141

## ⚙️ Variables de entorno

| Variable | Servicio | Descripción |
|----------|----------|-------------|
| `PythonApi__BaseUrl` | WebApp | URL de la Python API (default: `http://localhost:8000`) |
| `ANTHROPIC_API_KEY` | WebApp | Clave API de Anthropic para consultas LLM (opcional) |
| `RIMAC_EMAIL` / `RIMAC_PASSWORD` | WebApp | Credenciales SISGAQSAC para descarga automática (opcional) |
| `LP_USUARIO` / `LP_PASSWORD` | WebApp | Credenciales Agroevaluaciones para descarga automática (opcional) |
| `PORT` | Ambos | Puerto de escucha (WebApp: 10000, Python: 8000) |

## 🧭 Uso

1. Abrir la aplicación en el navegador
2. Subir los 2 archivos Excel: reporte MIDAGRI (La Positiva) y registro de siniestros (Rimac)
3. Hacer clic en "Procesar datos"
4. Navegar por las páginas del dashboard
5. Ir a "Reportes" para generar y descargar documentos

## 📦 Despliegue en producción

Ver [docs/GUIA_DESPLIEGUE_MIDAGRI.md](docs/GUIA_DESPLIEGUE_MIDAGRI.md) para instrucciones detalladas de despliegue en infraestructura MIDAGRI (Docker o IIS + Windows Service).

## 🛠️ Tecnologías

- .NET 8 / ASP.NET Core / Blazor Server
- Python 3.11 / FastAPI / Uvicorn
- python-docx, openpyxl, fpdf2, python-pptx, matplotlib, pandas, numpy
- ClosedXML (lectura Excel en .NET)
- Docker / Docker Compose
- Claude API (consultas LLM, opcional)
- Playwright (descarga automática, opcional)

## 📄 Licencia

MIT — ver [LICENSE](LICENSE).

## 👤 Autor

**Ivan Mena** — Director DSFFA, MIDAGRI · [@ivanmena2021](https://github.com/ivanmena2021) · [LinkedIn](https://www.linkedin.com/in/ivan-mena-r/)

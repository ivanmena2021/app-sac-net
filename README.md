# ReporteSAC — Sistema de Reportes del Seguro Agricola Catastrofico

Sistema web para el monitoreo y generacion de reportes del Seguro Agricola Catastrofico (SAC) 2025-2026, desarrollado para la Direccion General de Desarrollo Agricola — MIDAGRI.

## Arquitectura

El sistema consta de dos servicios que trabajan juntos:

```
Navegador (usuario)
    |
    v
WebApp (.NET 8 Blazor Server)    -- UI, carga de datos, logica de negocio
    |
    v  HTTP (multipart/form-data)
Python API (FastAPI)              -- Generacion de reportes Word, Excel, PDF, PPT
```

### WebApp (.NET 8)
- **Framework:** ASP.NET Core 8, Blazor Server (InteractiveServer)
- **Arquitectura:** Clean Architecture (Domain, Application, Infrastructure, WebApp)
- **Paginas:** Dashboard, Reportes, Departamental, Semaforo, Mapa de Calor, Calendario, Clima y Riesgo, Comparativo, Explorar, Consultas LLM

### Python API (FastAPI)
- **Framework:** FastAPI + Uvicorn
- **Generadores:** Word (python-docx), Excel (openpyxl), PDF (fpdf2), PowerPoint (python-pptx)
- **Datos estaticos:** Consolidados historicos, materia asegurada, primas, perfiles de riesgo

## Estructura del proyecto

```
ReporteSAC/
├── ReporteSAC.sln
├── Dockerfile                    # Build .NET app
├── docker-compose.yml            # Orquesta ambos servicios
├── docs/
│   └── GUIA_DESPLIEGUE_MIDAGRI.md
├── src/
│   ├── Domain/                   # Entidades del dominio (SiniestroRecord, etc.)
│   ├── Application/              # DTOs, interfaces de servicios, logica de procesamiento
│   │   ├── Contracts/Services/   # IWordReportService, IExcelReportService, etc.
│   │   ├── DTOs/                 # DatosNacionalesDto, DatosDepartamentalDto, etc.
│   │   └── Services/             # DataProcessorService
│   ├── Infrastructure/           # Implementaciones
│   │   ├── PythonApi/            # PythonApiReportService (HttpClient → Python API)
│   │   ├── Session/              # UploadedFileStore (scoped per-user)
│   │   ├── Alertas/              # SemaforoService
│   │   ├── AutoDownload/         # Descarga automatica desde portales aseguradores
│   │   ├── LlmQuery/             # Consultas en lenguaje natural (Claude API)
│   │   └── ExcelReader/          # Lectura de archivos Excel
│   └── WebApp/                   # Blazor Server app
│       ├── Components/Pages/     # 12 paginas Razor
│       ├── wwwroot/              # CSS, JS, iconos
│       └── Program.cs            # Entry point
└── python-api/
    ├── Dockerfile
    ├── main.py                   # FastAPI app con endpoints de generacion
    ├── requirements.txt
    ├── data_processor.py         # Procesamiento de datos para generadores
    ├── gen_word_nacional_py.py   # Ayuda memoria nacional
    ├── gen_word_departamental_py.py  # Ayuda memoria departamental
    ├── gen_word_operatividad.py  # Reporte de operatividad
    ├── gen_excel_eme.py          # Reporte EME
    ├── gen_excel_enhanced.py     # Excel mejorado con graficos
    ├── gen_pdf_resumen.py        # Resumen ejecutivo PDF
    ├── gen_ppt_dinamico.py       # PPT dinamica con filtros
    ├── gen_ppt_historico.py      # PPT historica por departamento
    └── static_data/              # Archivos estaticos requeridos
```

## Reportes generados

| Reporte | Formato | Descripcion |
|---------|---------|-------------|
| Ayuda Memoria Nacional | .docx | Resumen ejecutivo nacional con tablas y metricas |
| Ayuda Memoria Departamental | .docx | Detalle por departamento con provincias y eventos |
| Operatividad | .docx | Estado operativo del SAC por departamento |
| Reporte EME | .xlsx | Formato oficial para la Oficina de Emergencias |
| Excel Mejorado | .xlsx | Consolidado con graficos y multiples hojas |
| Resumen Ejecutivo | .pdf | Resumen visual con graficos y KPIs |
| PPT Dinamica | .pptx | Presentacion con filtros por empresa, zona, periodo |
| PPT Historica | .pptx | Comparativo historico por departamento |

## Inicio rapido

### Opcion 1: Docker Compose (recomendado)

```bash
git clone <repo-url>
cd ReporteSAC
docker compose up --build -d
```

Acceder a: http://localhost:5141

### Opcion 2: Desarrollo local

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

## Variables de entorno

| Variable | Servicio | Descripcion |
|----------|----------|-------------|
| `PythonApi__BaseUrl` | WebApp | URL de la Python API (default: `http://localhost:8000`) |
| `ANTHROPIC_API_KEY` | WebApp | Clave API de Anthropic para consultas LLM (opcional) |
| `RIMAC_EMAIL` / `RIMAC_PASSWORD` | WebApp | Credenciales SISGAQSAC para descarga automatica (opcional) |
| `LP_USUARIO` / `LP_PASSWORD` | WebApp | Credenciales Agroevaluaciones para descarga automatica (opcional) |
| `PORT` | Ambos | Puerto de escucha (WebApp: 10000, Python: 8000) |

## Uso

1. Abrir la aplicacion en el navegador
2. Subir los 2 archivos Excel: reporte MIDAGRI (La Positiva) y registro de siniestros (Rimac)
3. Hacer clic en "Procesar datos"
4. Navegar por las paginas del dashboard
5. Ir a "Reportes" para generar y descargar documentos

## Despliegue en produccion

Ver [docs/GUIA_DESPLIEGUE_MIDAGRI.md](docs/GUIA_DESPLIEGUE_MIDAGRI.md) para instrucciones detalladas de despliegue en infraestructura MIDAGRI (Docker o IIS + Windows Service).

## Tecnologias

- .NET 8 / ASP.NET Core / Blazor Server
- Python 3.11 / FastAPI / Uvicorn
- python-docx, openpyxl, fpdf2, python-pptx, matplotlib, pandas, numpy
- ClosedXML (lectura Excel en .NET)
- Docker / Docker Compose
- Claude API (consultas LLM, opcional)
- Playwright (descarga automatica, opcional)

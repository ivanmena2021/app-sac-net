# Guia de Despliegue — ReporteSAC MIDAGRI

**Version:** 1.0
**Fecha:** Abril 2026
**Destinatario:** Oficina de Tecnologias de la Informacion, MIDAGRI
**Clasificacion:** Uso interno

---

## 1. Vision general de la arquitectura

El sistema ReporteSAC consta de **dos servicios** que trabajan en conjunto:

| Servicio | Tecnologia | Funcion |
|----------|-----------|---------|
| **WebApp** | .NET 8 Blazor Server | Interfaz web, carga de archivos Excel, descarga de reportes, consultas LLM |
| **Python API** | Python 3.11, FastAPI | Generacion de reportes Word, Excel, PDF y PowerPoint a partir de los datos cargados |

### Diagrama de red

```
                         HTTPS (443)
  Navegador  ──────────────────────────►  IIS / Reverse Proxy
  (usuario)                                      │
                                                 │ HTTP (5141 / 10000)
                                                 ▼
                                        ┌─────────────────┐
                                        │   WebApp (.NET)  │
                                        │  Blazor Server   │
                                        │  Puerto 10000    │
                                        └────────┬────────┘
                                                 │
                                                 │ HTTP POST (multipart/form-data)
                                                 │ PythonApi__BaseUrl
                                                 ▼
                                        ┌─────────────────┐
                                        │  Python API      │
                                        │  FastAPI/Uvicorn │
                                        │  Puerto 8000     │
                                        └────────┬────────┘
                                                 │
                                                 │ HTTPS (salida)
                                                 ▼
                                          api.anthropic.com
                                        (solo si se usa LLM)
```

El usuario **nunca** se conecta directamente a la Python API; toda la comunicacion pasa por la WebApp .NET.

---

## 2. Prerrequisitos

### Opcion A — Docker (recomendado)

- Docker Engine 24+ o Docker Desktop
- Docker Compose v2
- Minimo 2 GB de RAM disponible
- Puerto 5141 libre (o el que se designe)

### Opcion B — IIS + Windows Service

- Windows Server 2019 o superior
- IIS 10 con los modulos:
  - ASP.NET Core Hosting Bundle 8.x
  - URL Rewrite
  - Application Request Routing (ARR) 3.0
  - WebSocket Protocol
- .NET 8 Runtime (ASP.NET Core Runtime)
- Python 3.11.x (instalacion de sistema, no Anaconda)
- NSSM (Non-Sucking Service Manager) para ejecutar la API Python como servicio Windows
- Acceso de red saliente a `api.anthropic.com` (puerto 443) si se habilita la funcion LLM

---

## 3. Opcion A: Despliegue con Docker Compose (mas simple)

### 3.1 Preparar el entorno

Copiar el repositorio completo al servidor. Verificar que existen:

```
ReporteSAC/
  Dockerfile
  docker-compose.yml
  ReporteSAC.sln
  src/
  python-api/
    Dockerfile
    main.py
    requirements.txt
    static_data/        <-- archivos estaticos requeridos
```

### 3.2 Configurar variables de entorno (opcional)

Si se desea habilitar la funcion de consulta LLM, crear un archivo `.env` en la raiz:

```env
ANTHROPIC_API_KEY=sk-ant-api03-xxxxxxxx
```

Y descomentar la linea correspondiente en `docker-compose.yml`.

### 3.3 Construir y levantar

```bash
cd ReporteSAC
docker compose up --build -d
```

### 3.4 Verificar

```bash
# Estado de los contenedores
docker compose ps

# Health check de la API Python
curl http://localhost:8000/health

# Acceder a la aplicacion web
# Abrir en navegador: http://localhost:5141
```

### 3.5 Detener

```bash
docker compose down
```

### 3.6 Actualizar

```bash
git pull
docker compose up --build -d
```

---

## 4. Opcion B: Despliegue con IIS + Windows Service

### 4.1 Publicar la aplicacion .NET

En una maquina con el SDK de .NET 8 instalado:

```powershell
cd ReporteSAC
dotnet publish src/WebApp/WebApp.csproj -c Release -o C:\publish\sac-webapp
```

Copiar la carpeta `C:\publish\sac-webapp` al servidor de produccion, por ejemplo a `C:\inetpub\sac-webapp`.

### 4.2 Configurar IIS para la WebApp

1. **Crear un Application Pool:**
   - Nombre: `SacWebAppPool`
   - .NET CLR Version: **No Managed Code**
   - Pipeline Mode: Integrated
   - Start Mode: AlwaysRunning
   - Idle Timeout (minutos): 0
   - Recycling > Regular Time Interval: 0 (deshabilitar reciclaje automatico)

2. **Crear un sitio o aplicacion IIS:**
   - Ruta fisica: `C:\inetpub\sac-webapp`
   - Binding: puerto designado (ej. 5141) o HTTPS 443 con certificado
   - Application Pool: `SacWebAppPool`

3. **Configurar `web.config`** (se genera automaticamente con `dotnet publish`, pero verificar):

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <location path="." inheritInChildApplications="false">
    <system.webServer>
      <handlers>
        <add name="aspNetCore" path="*" verb="*"
             modules="AspNetCoreModuleV2" resourceType="Unspecified" />
      </handlers>
      <aspNetCore processPath="dotnet"
                  arguments=".\WebApp.dll"
                  stdoutLogEnabled="true"
                  stdoutLogFile=".\logs\stdout"
                  hostingModel="InProcess">
        <environmentVariables>
          <environmentVariable name="ASPNETCORE_ENVIRONMENT" value="Production" />
          <environmentVariable name="PythonApi__BaseUrl" value="http://localhost:8000" />
        </environmentVariables>
      </aspNetCore>
    </system.webServer>
  </location>
</configuration>
```

### 4.3 Instalar la Python API como Windows Service con NSSM

1. **Instalar dependencias Python:**

```powershell
cd C:\servicios\sac-python-api
python -m pip install -r requirements.txt
```

2. **Probar manualmente:**

```powershell
python -m uvicorn main:app --host 0.0.0.0 --port 8000
# Verificar: curl http://localhost:8000/health
```

3. **Registrar como servicio con NSSM:**

```powershell
nssm install SacPythonApi "C:\Python311\python.exe"
nssm set SacPythonApi AppParameters "-m uvicorn main:app --host 0.0.0.0 --port 8000"
nssm set SacPythonApi AppDirectory "C:\servicios\sac-python-api"
nssm set SacPythonApi DisplayName "SAC Python API - Generador de Reportes"
nssm set SacPythonApi Description "Microservicio FastAPI para generacion de reportes SAC MIDAGRI"
nssm set SacPythonApi Start SERVICE_AUTO_START
nssm set SacPythonApi AppStdout "C:\servicios\sac-python-api\logs\service-stdout.log"
nssm set SacPythonApi AppStderr "C:\servicios\sac-python-api\logs\service-stderr.log"
nssm set SacPythonApi AppRotateFiles 1
nssm set SacPythonApi AppRotateBytes 10485760

nssm start SacPythonApi
```

4. **Verificar el servicio:**

```powershell
nssm status SacPythonApi
curl http://localhost:8000/health
```

### 4.4 Configuracion de IIS ARR para SignalR (Sticky Sessions)

Blazor Server utiliza SignalR con WebSockets. Si se usa un balanceador de carga o ARR como reverse proxy, se debe garantizar afinidad de sesion:

1. **Habilitar WebSockets en IIS:**
   - Server Manager > Add Roles and Features > Web Server > Application Development > WebSocket Protocol

2. **Habilitar ARR:**
   - IIS Manager > Server (nodo raiz) > Application Request Routing > Server Proxy Settings
   - Marcar "Enable proxy"

3. **Configurar afinidad en la Server Farm** (si aplica balanceo de carga):
   - Server Farm > Load Balance > seleccionar **Client affinity**
   - Algoritmo: Cookie-based

4. **Regla de reescritura para la WebApp** (si ARR actua como reverse proxy):

```xml
<rewrite>
  <rules>
    <rule name="ReverseProxyToBlazor" stopProcessing="true">
      <match url="(.*)" />
      <action type="Rewrite" url="http://localhost:5141/{R:1}" />
      <serverVariables>
        <set name="HTTP_X_FORWARDED_PROTO" value="https" />
      </serverVariables>
    </rule>
  </rules>
</rewrite>
```

### 4.5 Configuracion de Timeouts (5 minutos para generacion de reportes)

La generacion de reportes puede tomar varios minutos. Se deben ajustar los timeouts:

**En IIS (web.config de la WebApp):**

```xml
<aspNetCore ... requestTimeout="00:05:00">
```

**En ARR (si se usa como proxy):**

- IIS Manager > Application Request Routing > Server Proxy Settings
  - Time-out (seconds): **300**

**En la WebApp (.NET), el HttpClient ya tiene timeout configurado internamente.**

---

## 5. Tabla de variables de entorno

| Variable | Servicio | Requerida | Valor por defecto | Descripcion |
|----------|----------|-----------|-------------------|-------------|
| `ASPNETCORE_ENVIRONMENT` | WebApp | No | `Production` | Entorno de ejecucion .NET |
| `PORT` | WebApp | No | `10000` | Puerto interno del contenedor .NET |
| `PythonApi__BaseUrl` | WebApp | Si | `http://localhost:8000` | URL base de la Python API. En Docker: `http://python-api:8000` |
| `ANTHROPIC_API_KEY` | WebApp | No* | _(ninguno)_ | Clave API de Anthropic para consultas LLM. *Requerida solo si se habilita esa funcion |
| `PORT` | Python API | No | `8000` | Puerto en que escucha Uvicorn |

---

## 6. Archivos de datos estaticos requeridos

La Python API requiere la carpeta `python-api/static_data/` con los siguientes archivos. **Sin estos archivos, la generacion de reportes fallara.**

| Archivo | Descripcion |
|---------|-------------|
| `consolidado_sac_2024_2025.xlsx` | Consolidado historico de la campana SAC 2024-2025 |
| `Materia_Asegurada_SAC_2025-2026.xlsx` | Materia asegurada de la campana vigente |
| `Primas_Totales_SAC_2020-2026.xlsx` | Serie historica de primas (6 campanas) |
| `Resumen_SAC_2025-2026.xlsx` | Resumen ejecutivo de la campana vigente |
| `calendario_cultivos_historico.json` | Calendario de cultivos por departamento |
| `mapeo_siniestros.json` | Mapeo de codigos de siniestros |
| `perfil_riesgo_distrital.json` | Perfil de riesgo a nivel distrital |
| `resumen_campanas.json` | Resumen consolidado de campanas historicas |
| `resumen_departamental.json` | Resumen por departamento |
| `series_temporales.json` | Series temporales para graficos historicos |
| `METODOLOGIA_DATOS.md` | Documentacion de metodologia de procesamiento |

Estos archivos se incluyen en el repositorio y se copian automaticamente al contenedor durante el build de Docker. Para despliegue sin Docker, copiarlos manualmente a la carpeta de trabajo de la Python API.

---

## 7. Consideraciones de seguridad

### 7.1 HTTPS

- **Obligatorio** en produccion. Configurar certificado TLS en IIS o en el reverse proxy.
- En Docker, colocar un reverse proxy (Nginx, Traefik o el balanceador institucional) delante del puerto 5141.
- Blazor Server transmite estado de la interfaz por WebSocket; sin HTTPS, este trafico es vulnerable a interceptacion.

### 7.2 CORS

- La Python API actualmente permite todos los origenes (`allow_origins=["*"]`). En produccion, restringir a la IP o dominio de la WebApp.
- En despliegue Docker, la comunicacion entre servicios ocurre en la red interna `sac-network` y no necesita CORS; sin embargo, si se expone el puerto 8000 externamente, limitar los origenes.

### 7.3 Acceso saliente a Anthropic API

- Si se habilita la funcion de consulta LLM, el servidor debe permitir trafico HTTPS saliente hacia `api.anthropic.com` (puerto 443).
- La clave API (`ANTHROPIC_API_KEY`) debe tratarse como secreto. No incluirla en archivos versionados; usar variables de entorno o un gestor de secretos.
- Si la red institucional usa proxy de salida, configurar las variables `HTTP_PROXY` / `HTTPS_PROXY` en el contenedor o en el entorno del servicio.

### 7.4 Firewall

- Solo los puertos del reverse proxy (443) deben estar expuestos a la red de usuarios.
- Los puertos 5141, 8000 y 10000 deben ser accesibles unicamente desde localhost o la red interna del servidor.

---

## 8. Almacenamiento de archivos por sesion

Los archivos Excel subidos por cada usuario se almacenan en un servicio `IUploadedFileStore` registrado como **Scoped**. En Blazor Server, "Scoped" equivale a una instancia por circuito SignalR (es decir, por pestana del navegador). Esto significa que **cada usuario tiene su propia copia de los archivos** y no hay interferencia entre sesiones concurrentes.

**Nota sobre `SharedDatos`:** Los datos procesados (`Home.SharedDatos`) aun utilizan un campo `static` para compartir estado entre paginas dentro del mismo proceso. En el escenario tipico del equipo SAC (1-3 analistas trabajando con los mismos archivos), esto es funcional. Si en el futuro se requiere uso concurrente con datos distintos por usuario, se debera migrar `SharedDatos` a un servicio Scoped similar al `IUploadedFileStore`.

---

## 9. Procedimiento de verificacion post-despliegue

1. Acceder a la URL de la aplicacion en el navegador.
2. Verificar que la pagina de inicio carga correctamente (conexion SignalR activa).
3. Subir los archivos Excel de prueba (MIDAGRI + Siniestros).
4. Generar al menos un reporte de cada tipo (Word, Excel, PDF, PPT).
5. Verificar que los archivos descargados se abren correctamente.
6. Comprobar el health check de la Python API: `GET /health`.
7. Revisar los logs en busca de errores:
   - Docker: `docker compose logs webapp` y `docker compose logs python-api`
   - IIS: `C:\inetpub\sac-webapp\logs\stdout` y logs del servicio NSSM

---

*Documento elaborado para la transferencia tecnologica del sistema ReporteSAC al equipo de TI de MIDAGRI.*

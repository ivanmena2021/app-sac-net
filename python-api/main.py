"""
FastAPI microservice for SAC report generation.
Wraps the original Python generators from app_sac_github.
Called by the .NET Blazor Server app via HTTP.
"""

import io
import os
import sys
import traceback
import logging
from datetime import datetime
from fastapi import FastAPI, UploadFile, File, Query, HTTPException
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Add current directory to path for local imports
sys.path.insert(0, os.path.dirname(__file__))

# Set matplotlib backend to non-interactive before any other matplotlib import
import matplotlib
matplotlib.use("Agg")

app = FastAPI(
    title="SAC Report Generator API",
    description="Microservicio de generacion de reportes para el SAC 2025-2026",
    version="1.0.0",
)

# Allow CORS from any origin (the .NET app calls this)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
async def health():
    import pandas as pd
    return {
        "status": "ok",
        "service": "sac-report-generator",
        "timestamp": datetime.now().isoformat(),
        "python_version": sys.version,
        "pandas_version": pd.__version__,
        "static_data_exists": os.path.isdir(os.path.join(os.path.dirname(__file__), "static_data")),
        "static_files": os.listdir(os.path.join(os.path.dirname(__file__), "static_data")) if os.path.isdir(os.path.join(os.path.dirname(__file__), "static_data")) else [],
    }


@app.post("/api/test-process")
async def test_process(
    midagri: UploadFile = File(...),
    siniestros: UploadFile = File(...),
):
    """Test endpoint: processes data and returns diagnostic info without generating a document."""
    try:
        from data_processor import process_dynamic_data
        midagri_bytes = await midagri.read()
        siniestros_bytes = await siniestros.read()
        logger.info(f"Received midagri: {len(midagri_bytes)} bytes, siniestros: {len(siniestros_bytes)} bytes")

        datos = process_dynamic_data(
            io.BytesIO(midagri_bytes),
            io.BytesIO(siniestros_bytes)
        )
        return {
            "status": "ok",
            "midagri_bytes": len(midagri_bytes),
            "siniestros_bytes": len(siniestros_bytes),
            "midagri_shape": list(datos["midagri"].shape),
            "midagri_columns": list(datos["midagri"].columns),
            "materia_shape": list(datos["materia"].shape) if "materia" in datos else None,
            "materia_columns": list(datos["materia"].columns) if "materia" in datos else None,
            "total_avisos": int(datos.get("total_avisos", 0)),
            "departamentos": datos.get("departamentos_list", []),
        }
    except Exception as e:
        tb = traceback.format_exc()
        logger.error(f"Test process error: {tb}")
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}\n\nTraceback:\n{tb}")


@app.post("/api/process-and-generate")
async def process_and_generate(
    report_type: str = Query(..., description="Type: word-nacional, word-departamental, word-operatividad, excel-eme, excel-enhanced, pdf-ejecutivo, ppt-dinamico, ppt-historico"),
    departamento: str = Query(None, description="Department name (for departamental/historico reports)"),
    midagri: UploadFile = File(..., description="MIDAGRI Excel file (.xlsx)"),
    siniestros: UploadFile = File(..., description="Siniestros Excel file (.xlsx)"),
):
    """
    Unified endpoint: receives 2 Excel files, processes data, generates document.
    Returns the document bytes with appropriate content-type.
    """
    try:
        # Read uploaded files
        midagri_bytes = await midagri.read()
        siniestros_bytes = await siniestros.read()
        logger.info(f"Generating {report_type} | midagri={len(midagri_bytes)}B siniestros={len(siniestros_bytes)}B depto={departamento}")

        # Process data (same as data_processor.process_dynamic_data)
        from data_processor import process_dynamic_data, get_departamento_data, load_primas_historicas

        datos = process_dynamic_data(
            io.BytesIO(midagri_bytes),
            io.BytesIO(siniestros_bytes)
        )
        logger.info(f"Data processed: {datos['midagri'].shape[0]} rows, {len(datos.get('departamentos_list',[]))} deptos")

        # Generate document based on report_type
        doc_bytes = None
        filename = ""
        content_type = ""
        fecha_str = datetime.now().strftime("%d_%m_%Y")

        if report_type == "word-nacional":
            from gen_word_bridge_py import generate_nacional_docx
            doc_bytes = generate_nacional_docx(datos)
            filename = f"Ayuda_Memoria_Resumen_SAC_2025-2026_{fecha_str}.docx"
            content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        elif report_type == "word-departamental":
            if not departamento:
                raise HTTPException(status_code=400, detail="Parameter 'departamento' is required for this report type")
            from gen_word_bridge_py import generate_departamental_docx
            depto_data = get_departamento_data(datos, departamento)
            doc_bytes = generate_departamental_docx(depto_data)
            depto_title = departamento.strip().title()
            filename = f"Ayuda_Memoria_SAC_{depto_title}_{fecha_str}.docx"
            content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        elif report_type == "word-operatividad":
            from gen_word_operatividad import generate_operatividad_docx
            doc_bytes = generate_operatividad_docx(datos)
            filename = f"AM_Operatividad_SAC_{fecha_str}.docx"
            content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        elif report_type == "excel-eme":
            from gen_excel_eme import generate_reporte_eme
            doc_bytes = generate_reporte_eme(datos)
            filename = f"formato_reporte_EME_{fecha_str}.xlsx"
            content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        elif report_type == "excel-enhanced":
            from gen_excel_enhanced import generate_enhanced_excel
            doc_bytes = generate_enhanced_excel(datos)
            filename = f"Consolidado_SAC_2025-2026_{fecha_str}.xlsx"
            content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        elif report_type == "pdf-ejecutivo":
            from gen_pdf_resumen import generate_executive_pdf
            doc_bytes = generate_executive_pdf(datos)
            filename = f"Resumen_Ejecutivo_SAC_{fecha_str}.pdf"
            content_type = "application/pdf"

        elif report_type == "ppt-dinamico":
            from gen_ppt_dinamico import generar_ppt_dinamico
            # Build filtros dict (no filters = national scope)
            filtros = {
                "empresa": "ambas",
                "tipos_siniestro": [],
                "fecha_inicio": None,
                "fecha_fin": None,
                "departamentos": [departamento] if departamento else [],
                "provincias": [],
                "distritos": [],
            }
            doc_bytes = generar_ppt_dinamico(
                datos["midagri"],
                filtros,
                datos["fecha_corte"]
            )
            scope = departamento.strip().title() if departamento else "Nacional"
            filename = f"SAC_{scope}_{fecha_str}.pptx"
            content_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

        elif report_type == "ppt-historico":
            if not departamento:
                raise HTTPException(status_code=400, detail="Parameter 'departamento' is required for this report type")
            from gen_ppt_historico import generar_ppt_historico
            primas_hist = load_primas_historicas()
            doc_bytes = generar_ppt_historico(
                departamento.strip().upper(),
                datos,
                primas_hist
            )
            depto_title = departamento.strip().title()
            filename = f"Historico_SAC_{depto_title}_{fecha_str}.pptx"
            content_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

        else:
            raise HTTPException(status_code=400, detail=f"Unknown report_type: {report_type}")

        if doc_bytes is None:
            raise HTTPException(status_code=500, detail="Document generation returned empty result")

        return Response(
            content=doc_bytes,
            media_type=content_type,
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"',
                "X-Filename": filename,
            },
        )

    except HTTPException:
        raise
    except Exception as e:
        tb = traceback.format_exc()
        logger.error(f"Error generating {report_type}: {tb}")
        raise HTTPException(status_code=500, detail=f"Error generating {report_type}: {str(e)}\n\nTraceback:\n{tb}")


@app.post("/api/process-data")
async def process_data_only(
    midagri: UploadFile = File(...),
    siniestros: UploadFile = File(...),
):
    """
    Process data only (no document generation). Returns JSON metrics.
    Used by .NET app to get processed data for UI display.
    """
    try:
        from data_processor import process_dynamic_data

        midagri_bytes = await midagri.read()
        siniestros_bytes = await siniestros.read()

        datos = process_dynamic_data(
            io.BytesIO(midagri_bytes),
            io.BytesIO(siniestros_bytes)
        )

        # Return serializable metrics (no DataFrames)
        return {
            "fecha_corte": datos["fecha_corte"],
            "total_avisos": int(datos["total_avisos"]),
            "total_ajustados": int(datos["total_ajustados"]),
            "pct_ajustados": round(datos["pct_ajustados"], 2),
            "ha_indemnizadas": round(datos["ha_indemnizadas"], 2),
            "monto_indemnizado": round(datos["monto_indemnizado"], 2),
            "monto_desembolsado": round(datos["monto_desembolsado"], 2),
            "productores_desembolso": int(datos["productores_desembolso"]),
            "prima_total": round(datos["prima_total"], 2),
            "prima_neta": round(datos["prima_neta"], 2),
            "sup_asegurada": round(datos["sup_asegurada"], 2),
            "prod_asegurados": int(datos["prod_asegurados"]),
            "indice_siniestralidad": round(datos["indice_siniestralidad"], 2),
            "pct_desembolso": round(datos["pct_desembolso"], 2),
            "deptos_con_desembolso": int(datos["deptos_con_desembolso"]),
            "departamentos_list": datos["departamentos_list"],
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing data: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)

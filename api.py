from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from caculadora import process_and_create_excel
from helpers import is_valid_file_path, clean_old_files
import logging
import shutil
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

ALLOWED_MIME_TYPES = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]

def validate_file_type(file: UploadFile):
    if file.content_type not in ALLOWED_MIME_TYPES:
        logging.warning("validate_file_type: Tipo de archivo no permitido.")
        raise HTTPException(status_code=415, detail="Tipo de archivo no permitido.")

# Llamar a la función de limpieza cuando la aplicación inicie
@app.on_event("startup")
async def startup_event():
    clean_old_files("C:\\caluladora\\pensiones")

@app.get("/health")
def health():
    return {"health": "OK"}

@app.post("/procesar")
async def procesar(
    process_type: str = Form(...),
    discount_percent: float = Form(...),
    money_formula: str = Form(None),
    payment_period: int = Form(None),
    retroactive_period: int = Form(None),
    file: UploadFile = File(...),
):
    
    validate_file_type(file)

    temp_file_path = f"temp_{file.filename}"
    generated_file = None

    try:
        # Guardar archivo subido temporalmente
        with open(temp_file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        if retroactive_period is None:
            retroactive_period = 0

        # Procesar el archivo
        generated_file = process_and_create_excel(
            process_type,
            discount_percent,
            money_formula,
            payment_period,
            retroactive_period,
            temp_file_path
        )

        # Verificación adicional
        if generated_file and os.path.exists(generated_file):
            logging.info(f"Endpoint: procesar, file: {generated_file}")
            return {"file_path": generated_file.split('/')[3]}
        else:
            logging.error(f"status: 500, Error al generar el archivo procesado, Endpoint: procesar")
            raise HTTPException(status_code=500, detail="Error al generar el archivo procesado, Endpoint: procesar")

    except Exception as e:
        logging.critical(f"status: 500, Endpoint: procesar, Error interno: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error interno: {str(e)}")
    finally:
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

@app.get("/descargar/{file_name}")  # :path permite rutas con directorios
async def descargar(file_name: str):
    path = f"C:/caluladora/pensiones/{file_name}"

    # Validar la ruta del archivo
    if not is_valid_file_path(path):
        logging.warning(f"Ruta no válida: {path}, Endpoint: descargar")
        raise HTTPException(status_code=406, detail="Ruta no válida")
    
    try:
        # Comprobar si el archivo existe en la ruta proporcionada
        if os.path.exists(path):
            logging.info(f"Endpoint: descargar, file: {path}")
            return FileResponse(path, filename=os.path.basename(path))
        else:
            logging.warning(f"Endpoint: descargar, Archivo no encontrado: {path}")
            raise HTTPException(status_code=404, detail="Archivo no encontrado.")
    except Exception as e:
        logging.critical(f"Endpoint: descargar, Error al intentar descargar el archivo: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error al intentar descargar el archivo: {str(e)}")

LOG_FILE_PATH = "calculadora.log"
@app.get("/descargar-log")
async def descargar_log():
    # Verificar si el archivo de log existe
    if not os.path.exists(LOG_FILE_PATH):
        logging.warning(f"Endpoint: descargar-log, Archivo de log no encontrado.")
        raise HTTPException(status_code=404, detail="Archivo de log no encontrado.")
    
    # Enviar el archivo como respuesta para descargarlo
    return FileResponse(
        path=LOG_FILE_PATH, 
        filename="calculadora.log", 
        media_type='text/plain'
    )
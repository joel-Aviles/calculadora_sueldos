### Documentación Técnica del Proyecto

#### 1. **Resumen del Proyecto**
Este proyecto es una aplicación basada en Python que implementa una calculadora con funcionalidad extendida. Incluye un sistema de registro de eventos (logs), una API para interactuar con la calculadora, y un conjunto de funciones auxiliares para soportar la lógica principal.

---

#### 2. **Estructura del Proyecto**
La estructura del proyecto está organizada de la siguiente manera:

```
calculadora/
├── api.py               # Lógica de la API para interactuar con la calculadora
├── calculadora.py       # Lógica principal de la calculadora
├── helpers.py           # Funciones auxiliares para soporte
├── calculadora.log      # Archivo de log para seguimiento de eventos
├── requirements.txt     # Dependencias del proyecto
├── TODO.md              # Lista de tareas pendientes
├── excel_files/         # Archivos necesarios para el buen funcionamineto de la calculadora
├── .git/                # Archivos de control de versiones con Git
└── .gitignore           # Archivos ignorados por Git
```

---

#### 3. **Descripción de Archivos Clave**

##### **Archivo Principal**
- **`calculadora.py`**: Contiene la implementación principal de la calculadora, incluyendo las operaciones matemáticas.

##### **API**
- **`api.py`**: Proporciona una interfaz para interactuar con la calculadora a través de una API RESTful. Este archivo permite realizar operaciones mediante solicitudes HTTP.

##### **Funciones Auxiliares**
- **`helpers.py`**: Incluye funciones auxiliares que soportan la lógica principal de la calculadora, como validación de entradas y manejo de errores.

##### **Archivo de Registro**
- **`calculadora.log`**: Registra eventos y errores para facilitar el depurado y el monitoreo del sistema.

##### **Dependencias**
- **`requirements.txt`**: Lista de bibliotecas necesarias para ejecutar el proyecto, incluyendo las versiones recomendadas.

---

#### 4. **Dependencias del Proyecto**
Este proyecto depende de las siguientes bibliotecas (listadas en `requirements.txt`):
- **FastAPI**: Para manejar la API RESTful.
- **Logging**: Biblioteca de Python para registro de eventos (parte del estándar).
- **Openpyxl**: Para crear y editar archivos de excel.
- **Pandas**: Para la manipulación de los datos dentro de excel

---

#### 5. **Guía de Instalación**
1. Clona el repositorio en tu máquina local:
   ```bash
   git clone https://github.com/joel-Aviles/calculadora_sueldos.git
   ```
2. Accede al directorio del proyecto:
   ```bash
   cd calculadora_sueldos
   ```
3. Crea un entorno virtual:
   ```bash
   python -m venv venv
   source venv/bin/activate  # En Windows: venv\Scripts\activate
   ```
4. Instala las dependencias:
   ```bash
   pip install -r requirements.txt
   ```
5. Ejecuta la aplicación:
   ```bash
   uvicorn api:app --reload
   ```

---

#### 6. **Documentación de endpoints**
- https://calculadora-sueldos.onrender.com/docs

---

#### 7. **Flujo de Datos**
- **Entrada de Usuario:** Los datos ingresados a través de la API son procesados por `api.py`.
- **Lógica Principal:** `calculadora.py` realiza las operaciones matemáticas y valida las entradas mediante funciones auxiliares de `helpers.py`.
- **Salida:** Los resultados son enviados de vuelta al usuario a través de respuestas HTTP.
- **Registro:** Los eventos y errores se almacenan en `calculadora.log`.

---

#### 8. **Prácticas de Desarrollo**
- **Modularidad:** Cada funcionalidad está separada en módulos específicos.
- **Registro de Eventos:** El sistema de logs facilita el seguimiento de errores y eventos importantes.
- **Control de Versiones:** Uso de Git para gestionar cambios en el código fuente.

---

#### 9. **Pendientes y Recomendaciones**
- **Documentación de la API:** Agregar ejemplos de solicitudes y respuestas para los endpoints disponibles.
- **Pruebas Unitarias:** Implementar pruebas para validar las operaciones de `calculadora.py` y las respuestas de `api.py`.
- **Optimizaciones:** Revisar el rendimiento del código en escenarios de alta carga.

---

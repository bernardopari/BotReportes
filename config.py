# config.py

# Fuente de datos: 'excel' o 'bd'
FUENTE_DATOS = 'excel'

# Excel
ARCHIVO_EXCEL = 'BD/datos_reporte.xlsx'

# Base de datos (si eliges 'bd')
DB_PATH = 'bd/datos.db'
DB_QUERY = 'SELECT * FROM ventas;'

# Correo
ARCHIVO = 'datos_reporte.xlsx'
DESTINATARIO = 'bparic@est.unap.edu.pe'
REMITENTE = 'bernardpari2002@gmail.com'
CONTRASENA = 'ehyk ipwy qrqy ixtw'  # Usa una app password si es Gmail

# Intervalo de tiempo (en minutos)
TIEMPO_REPORTE_MINUTOS = 2

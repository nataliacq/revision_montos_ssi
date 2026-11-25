considera 

- para madar mail de notificacion he establecido contraseña de aplicación. la guarde con comando setx GMAIL_APP_PASSWORD "***".

- libro donde esta proyeccion26 https://docs.google.com/spreadsheets/d/1u4JHmEyIn6WusJWZLF6DQPDyhGxgRF5cEEp4z_oE8iw/edit?gid=1133143073#gid=1133143073 .

- "C:\Users\HP\OneDrive\daily-new-inv\descarga_diaria_mef_23_11_26_3pm.gs" script google apps que escribe datos en google sheet file drive_exclusivo.

- "C:\Users\HP\OneDrive\daily-new-inv\credenciales" para modificar drive_exclusivo.

**objetivo** crea el script de python. pasos:

1. obtener data: 
   
   1. columna A de hoja PIM-diario desde A3 hacia abajo tiene cui  recien_insertados. 
   
   2. todos_cuis: columna A de hoja Proyeccion26 (su encabezado dice CUI en A2) cui que tengan como etapa cualquier valor (incluso vacio) menos "PC". etapa esta en hoja proyeccion26 columna F, F2 tiene encabezado "Etapa" y a partir de F3 estan los CUI. 

2. crea txt con json o formato que consideres con recien_insertados y todos_cuis.

3. NOTA: 
- si algun nombre de encabezado no coincide con texto que esta en columna que se señala en que debe estar, envía mail aviso detallado y parar ejecucion.
- considera casos "a partir de F3 estan los CUI" o "desde A3 hacia abajo tiene los cui recien_insertados". en casos como estos, se debe verificar que realmente empieza CUI desde aca hacia abajo pues alguien pudo insertar fila arriba con otro texto. y una vez copiado cuis de la hoja de google debes validar que cada uno son valores numericos de 7 cifras pues alguien pudo escribir por error el una celda, si es este caso, avisar por correo.



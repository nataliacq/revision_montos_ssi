# HANDOFF – Montos de inversión, UEI y tipo de inversión desde el SSI del MEF

Última actualización: 2026-11-25 (15:40)
Sesión de Claude Code: "monto inversion ssi"

## Objetivo
Cada día, para los CUI del libro de Google (hojas Proyección26 y PIM-diario), traer del SSI del MEF
(https://ofi5.mef.gob.pe/ssi/) y escribir en PIM-diario:
- para todos los CUI que no son PC: el monto de inversión = COSTO TOTAL DE LA INVERSIÓN ACTUALIZADO (S/) (a+b+c+d);
- para los CUI recién insertados, además: UNIDAD EJECUTORA DE INVERSIONES (UEI) y TIPO DE INVERSIÓN.

## Estado en una línea
**Programa de Python terminado y probado contra el libro real** (prompts 1, 2 y 3 de `nata_apuntes`).
Falta programarlo a las 6:45 en el Programador de tareas de Windows (no pedido aún) y las mejoras de
`prompt_xall.md`.

## 1. Archivos
| Qué | Dónde |
|---|---|
| Programa | `4-script\cuis_desde_sheets.py` (≈800 líneas, todo en un archivo) |
| Pruebas sin red | `4-script\pruebas_sin_red.py` (simulan hoja, corridas anteriores y correo; `python pruebas_sin_red.py`) |
| Excel de variaciones | `...\monto_ssi\variacion_montos_AAAA-MM-DD_HHMMSS.xlsx` (solo si cambió algún monto) |
| Pedidos de la usuaria | `nata_apuntes\prompt1.md`, `prompt2.md`, `prompt3.md` (hechos); `prompt_xall.md` (ideas sin pedir) |
| Listas de CUI por corrida | `nata_apuntes\test_responses\cuis_AAAA-MM-DD_HHMMSS.json` |
| Montos por corrida | `C:\Users\HP\Documents\drives_exclusivos_descarga\monto_ssi\monto_ssi_AAAA-MM-DD_HHMMSS.json` y `.csv` (separador `;`, coma decimal) |
| Credencial de Google | `credenciales\credenciales.json` (cuenta de servicio, ya compartida como editora del libro: escribió bien) |
| Avisos no enviados | `nata_apuntes\test_responses\AVISO_NO_ENVIADO_*.txt` (solo si falla el correo) |
| Guion de Apps Script vigente | `descarga_diaria_mef_23_11_26_3pm.gs` (llena PIM-diario C:L, A1, A3:A y N2:N) |
| Apps Script de montos (abandonado) | `4-script\montos_ssi.gs`: el SSI bloquea a Google, ver sección 5 |
| Excel puntual del 12/11 | `C:\Users\HP\OneDrive\ssi monto inv\2 cuis-montos.xlsx` (359 CUI, montos + devengado 2025) |

## 2. Cómo se usa
```
cd C:\Users\HP\OneDrive\daily-new-inv\4-script
python cuis_desde_sheets.py                 # corrida completa (~10 s)
python cuis_desde_sheets.py --probar-correo # solo manda un correo de prueba
```
Códigos de salida: 0 bien (aunque haya avisos enviados); 2 detenido por títulos o estructura;
3 error inesperado; 4 no se pudo enviar el correo; 5 el SSI no respondió para ningún CUI.
Requiere Python con `gspread` (6.2.1 instalado) y `curl.exe` de Windows.

## 3. Qué hace (3 pasos)
Configuración al inicio del programa: `ID_LIBRO`, nombres de hojas, `CORREO_AVISO`, carpetas,
posiciones de títulos, `CUIS_SIN_AVISO`.

### Paso 1 – leer las listas (una sola lectura del libro)
- `recien_insertados` = PIM-diario, columna A desde A3.
- `todos_cuis` = Proyección26, columna A desde A3, solo filas con Etapa (columna F) distinta de "PC"
  (Etapa vacía sí entra; "pc " con espacios o minúscula cuenta como PC).
- Controles que DETIENEN sin escribir nada (con correo): Proyección26 A2 ≠ "CUI" o F2 ≠ "Etapa";
  PIM-diario C1 ≠ "CUI" o A2 no vacía. La columna A de PIM-diario no tiene título propio, por eso
  C1 y A2 sirven de ancla para detectar filas insertadas o borradas arriba.
  La comparación de títulos ignora mayúsculas, tildes y espacios.
- Valores que no son número de 7 cifras: se dejan fuera, se anotan con su celda y se avisan; el resto sigue.
- Repetidos: se omiten (anotados en el archivo).
- La hoja se busca por nombre exacto y, si no, sin tildes (Proyeccion26 = Proyección26).

### Paso 2 – descargar del SSI
- Servicio `traeDetInvSSI` (el de la pantalla del buscador), POST `id=<CUI>`, con `curl.exe`,
  12 consultas en paralelo, 3 intentos por CUI. 229 CUI en ~5 s.
- Se consulta la unión de las dos listas (si un insertado no está en todos_cuis, se consulta igual y se avisa).
- Monto = `MTO_F8_C + MTO_LAUDO + MTO_CartFza` (misma fórmula de la página). Monto vacío -> 0.
- UEI = `DES_UNIDAD_UEI`; tipo = `TIPO_FORMATO`.
- CUI que el SSI no encuentra (lista vacía) o que no responde: monto null (no 0) y aviso.
- `CUIS_SIN_AVISO = {3000882, 3999999}`: no existen en el SSI y la usuaria pidió no avisarlos
  (siguen anotados en los archivos).
- "NO REGISTRADO EN EL INVIERTE" (hoy 6, p. ej. los genéricos 2000270, 2000351, 2001621, 2016766):
  monto 0, sin aviso (si no, llegaría un correo cada día), marcados en el archivo.

### Paso 3 – escribir en PIM-diario (una lectura y una escritura, solo O, P y Q)
- Tabla de N a Q, **una fila por CUI de la columna N** (`CUI_proyeccion26_noPC`), ubicada por CUI,
  no por posición. Decisión de diseño de Claude (la usuaria no la objetó; confirmar si se cambia):
  - O `COSTO_TOTAL_INVERSION_ACTUALIZADO (a+b+c+d)` = monto de ese CUI, formato `#,##0.00`.
    Si el monto es null, se deja el valor que ya tenía la celda (no se pisa con 0).
  - P `UEI` y Q `TIPO_INVERSION` = solo en filas de recién insertados; las demás quedan vacías.
- Título de O, P o Q que no coincide: esa columna no se escribe, las otras sí, con aviso.
  N1 que no coincide: no se escribe nada, con aviso.
- Filas sobrantes de corridas anteriores (si N se acortó) se limpian.
- Si el SSI se cayó entero, no se escribe nada.
- Recién insertado que no está en N: queda sin UEI ni tipo y se avisa.

### Comparación con la corrida anterior (agregado 25/11 tarde)
- Base = último `monto_ssi_*.json` de la carpeta de montos (la corrida anterior, aunque sea una prueba del
  mismo día). Compara su `monto_final` (o `monto_inversion` en archivos viejos) con lo descargado hoy.
- Montos que cambiaron -> `variacion_montos_AAAA-MM-DD_HHMMSS.xlsx` en la misma carpeta (solo si hubo
  cambios): CUI, nombre de la inversión (`NOMBRE_INVERSION`, ahora se descarga), monto anterior, monto
  actual, % de variación con signo (negativo = bajó). Filas resaltadas si pasan el umbral o si antes era 0
  ("antes 0"). Hoja aparte "Ceros sospechosos" y hoja "Nota".
- Constante global `UMBRAL_VARIACION_PCT = 10`: el correo alerta solo variaciones mayores a ese % (en más o
  en menos); las demás solo van al Excel.
- Dato erróneo: si un CUI llega con monto 0 y en la corrida anterior tenía monto > 0, se mantiene el
  anterior (en la hoja y como `monto_final` del archivo) y se avisa **solo la primera vez**. Si en las
  corridas siguientes sigue llegando en 0, se sigue manteniendo el anterior **sin alerta** (pedido de la
  usuaria: no repetir el aviso a diario); queda anotado en el campo `observacion` del archivo. Se detecta
  porque la corrida anterior tiene `monto_inversion` = 0 y `monto_final` > 0. Si el SSI no da dato,
  también se mantiene el anterior. El archivo guarda ambos: `monto_inversion` (descargado) y `monto_final` (usado).
- Probado con corridas anteriores simuladas en `pruebas_sin_red.py` (alza 5 %, baja 15 %, caída a 0,
  0 repetido el día 2 sin alerta, recuperación el día 3, sin dato, antes 0, igual, nuevo) y en real a
  las 15:24 y 15:34 (0 cambios, sin Excel ni correo, como corresponde).

### Correo
- Un solo correo por corrida con todos los avisos (asunto `[CUI PIM-diario] ...`), por Gmail SMTP
  (smtp.gmail.com:465). Usuario y destinatario = `CORREO_AVISO` (el primero envía).
- Contraseña: variable de usuario de Windows `GMAIL_APP_PASSWORD` (creada con `setx`); si la ventana
  es anterior al `setx`, el programa la lee del registro (HKCU\Environment).
- Si el correo falla, el aviso se guarda en `AVISO_NO_ENVIADO_*.txt`.

## 4. Pruebas hechas (25/11)
- Listas: coinciden exactamente con lo que el Apps Script deja en A3:A (13) y N2:N (229 = 348 con CUI − 119 PC).
- Casos simulados en memoria (sin tocar la hoja): títulos movidos, fila insertada arriba en cada hoja,
  valores inválidos, repetidos, Etapa "pc ", servicio caído, CUI inexistente, P1 cambiado, N1 cambiado,
  monto null que conserva el anterior, filas sobrantes.
- Montos: 223 de 227 iguales al Excel del 12/11; 4 subieron por actualización del MEF
  (2195115, 2307889, 2482057, 2546561).
- Corrida real 12:21: 229 filas escritas, releídas y verificadas (227 montos, 2 vacíos esperados,
  13 insertados con UEI y tipo, 0 errores). Sin correo (no hubo avisos).
- Correo de prueba enviado y recibido el 25/11.
- La usuaria probó a mano poniendo "aaa" en Proyección26 A5 y A17 (corrida 10:58): se detectó bien.
  Revisar que esas celdas hayan vuelto a su CUI (en la corrida de 12:21 ya no había inválidos).

## 5. Contexto técnico importante
- **ofi5.mef.gob.pe bloquea a los servidores de Google** (diagnóstico 23/11: https, sin verificar
  certificado y http, todos "Dirección no disponible" tras 50 s; apps5.mineco sí responde). Por eso
  esto corre en la PC y no en Apps Script.
- **Certificado del MEF incompleto**: Python `requests` falla con CERTIFICATE_VERIFY_FAILED. No
  desactivar la verificación; `curl.exe` de Windows completa la cadena. Alternativa limpia: `truststore`.
- `python -I` ignora los paquetes del usuario (gspread, openpyxl no aparecen).
- Fuente alternativa no validada: `https://fs.datosabiertos.mef.gob.pe/datastorefiles/DETALLE_INVERSIONES.csv`
  (238 MB, 68 columnas: COSTO_ACTUALIZADO, NOMBRE_UEI, TIPO_INVERSION, MONTO_ET_F8, MONTO_LAUDO, MONTO_FIANZA...).
- El MEF escribe algunas UEI en minúsculas (2106687, 2195115: "gerencia de infraestructura del proyecto
  especial..."); se dejan tal cual. Pasar a mayúsculas es un cambio de una línea si se pide.
- Receta del devengado por año (ícono btn_finan): POST `/invierteWS/Dashboard/traeDevengSSI` con
  `id=<CUI>&tipo=DEV2`; sumar orígenes del mismo año.

## 6. Para la revisión (decisiones y riesgos que conviene mirar con lupa)
Decisiones tomadas por Claude, no pedidas literalmente por la usuaria:
1. **Alineación de O, P y Q por el CUI de la columna N** (tabla N:Q). Alternativa descartada: alinear P y Q
   con la columna A. Riesgo: si N está desactualizada respecto de Proyección26 (la llena el Apps Script en
   otro momento), algunos CUI descargados no tienen fila; hoy se avisa.
2. **Monto null conserva el valor anterior** en vez de poner 0 (para no pisar datos buenos si el MEF falla).
3. **Cero sospechoso = de > 0 a 0.** Riesgo conocido: si un costo de verdad pasa a 0, la hoja conservará el
   valor viejo indefinidamente y solo se avisa el primer día. Un cero en un CUI que antes era 0 (p. ej. los
   "no registrados") no es sospechoso.
4. **Base de comparación = el último archivo de la carpeta**, aunque sea una prueba del mismo día. No hay
   distinción entre corridas de prueba y la diaria.
5. **Los "NO REGISTRADO EN EL INVIERTE" no avisan** (monto 0 legítimo, hoy 6).
6. **Un solo correo por corrida**; si el correo falla, el aviso queda en un .txt.

Puntos técnicos a revisar:
- Lectura de la hoja con `UNFORMATTED_VALUE` y validación de CUI (`como_cui`: int, float entero o texto de
  7 dígitos).
- Escritura con `valueInputOption RAW` de O2:Q(n) en una sola llamada; limpieza de filas sobrantes.
- `consultar_ssi`: clasifica respuesta (lista vacía = no existe; error de red, HTTP ≠ 200 o no-JSON = falla),
  3 intentos con pausa. Depende de `curl.exe` (solo Windows).
- No hay bloqueo contra dos corridas simultáneas (p. ej. la tarea programada y una prueba manual).
- La contraseña del correo se lee de la variable `GMAIL_APP_PASSWORD` o del registro de Windows.
- `pruebas_sin_red.py` no cubre la consulta real al SSI, la lectura real del libro ni el envío del correo.

## 7. Pendientes
1. **Tarea diaria a las 6:45** (Programador de tareas de Windows). Debe correr después de la
   descarga de Apps Script (6:20) y solo corre si la PC está prendida. Usar la ruta completa de
   python.exe y "Iniciar en" = carpeta `4-script`. Considerar "ejecutar lo antes posible si se perdió".
2. **Confirmar el libro**: `ID_LIBRO` apunta a "Copia de Drives exclusivos 19.11". Si el definitivo es
   otro, cambiar la constante y compartir ese libro con la cuenta de servicio (correo en
   `credenciales\correo iam.txt`).
3. **Seguridad**: la contraseña de aplicación de Gmail está en texto plano en
   `nata_apuntes\apuntes ssi monto script.txt` y `nata_apuntes\mail_notif_daily_script.md` (OneDrive) y
   pasó por el chat. Recomendado: borrarla de esos archivos, generar una nueva y volver a `setx`.
4. Ideas en `prompt_xall.md`, no pedidas aún (la comparación por CUI con la corrida anterior ya está
   hecha): revisar devengado acumulado exorbitante, ver qué fórmulas dependen de
   las columnas que se modifican, "si no está la columna Etapa, buscarla en la fila" (hoy se detiene).
5. Corregir o quitar de Proyección26 los CUI 3000882 y 3999999 (no existen en el SSI).
6. `4-script\montos_ssi.gs` ya no sirve (bloqueo); se puede borrar del proyecto de Apps Script.

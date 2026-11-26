# HANDOFF – Montos de inversión, UEI y tipo de inversión desde el SSI del MEF

Última actualización: 2026-11-26
Sesión de Claude Code: "monto inversion ssi" (25/11) + revisión y cambios del 26/11

## Objetivo
Cada día, para los CUI del libro de Google (hojas Proyección26 y PIM-diario), traer del SSI del MEF
(https://ofi5.mef.gob.pe/ssi/) y escribir **en Proyección26**, ubicando cada fila por su CUI (columna A):
- **I** "(A) Monto de inversión" = COSTO TOTAL DE LA INVERSIÓN ACTUALIZADO (S/) (a+b+c+d), para
  **todos** los CUI de la lista (los que no son PC);
- **D** "UEI" = la sigla equivalente a la UEI del SSI (GRI, GRDE, GRDIS, GRRNyGA, PET, DIRESA, DRAT,
  DRTYC), **solo en las filas de los recién insertados**;
- **F** "Etapa" = "PC" si el SSI dice PROYECTO PROCOMPITE, "Desactivado" si está desactivado,
  **solo en las filas de los recién insertados**.

Cambio del 26/11: antes se escribía en PIM-diario (columnas O, P y Q) para después llevarlo a
Proyección26 con fórmulas. Se escribe directo en Proyección26 porque PIM-diario se rearma cada día
y los valores de los insertados del día anterior se perdían. Las columnas O, P y Q de PIM-diario
ya no se usan y se vacían.

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
| Guion de Apps Script vigente | `descarga_diaria_mef_23_11_26_3pm.gs` (llena PIM-diario C:L, A1, A3:A y N2:N, y vacía O:Q) |
| Apps Script de montos (abandonado) | `4-script\montos_ssi.gs`: el SSI bloquea a Google, ver sección 5 |
| Excel puntual del 12/11 | `C:\Users\HP\OneDrive\ssi monto inv\2 cuis-montos.xlsx` (359 CUI, montos + devengado 2025) |

## 2. Cómo se usa
```
cd C:\Users\HP\OneDrive\daily-new-inv\4-script
python cuis_desde_sheets.py                 # corrida completa (~10 s)
python cuis_desde_sheets.py --probar-correo # solo manda un correo de prueba
```
Códigos de salida: 0 bien (aunque haya avisos enviados); 2 detenido por títulos o estructura;
3 error inesperado; 4 no se pudo enviar el correo; 5 el SSI no respondió para ningún CUI;
6 detenido por ceros masivos (MAX_CEROS_SOSPECHOSOS, data basura del SSI).
Solo corre una instancia a la vez (`nata_apuntes\corrida_en_curso.lock`): si la tarea programada y
una prueba manual coinciden, la segunda no hace nada.
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

### Paso 3 – escribir en Proyección26 (ubicando cada fila por su CUI de la columna A)
- **I** `(A) Monto de inversión`, en **todas** las filas de la lista (los CUI que no son PC): si el
  SSI no dio monto, la celda **no se toca** (nunca se pisa un dato bueno con 0 ni con vacío). No se
  le cambia el formato de número: es una columna de la hoja.
- **D** `UEI` y **F** `Etapa`, **solo en las filas de los recién insertados**. Esto hace que la
  garantía de no pisar nada sea **estructural** y no una comparación de valores: el guion de Google,
  al insertar una fila en Proyección26, escribe únicamente A y B y copia formato y fórmulas de
  M:AA, así que D y F llegan vacías. En una fila que ya estaba, esas dos celdas no se tocan nunca.
- **D** `UEI`: la sigla equivalente (tabla `UEI_EQUIVALENCIAS`). Si el texto del SSI no coincide con
  ninguna, la celda **no se toca** y se avisa con el texto que llegó, para agregar la equivalencia.
- **F** `Etapa`: "PC" si `TIPO_FORMATO` es PROYECTO PROCOMPITE; "Desactivado" si `ESTADO` empieza
  con DESACTIVADO. Las demás etiquetas (ET, EF, L...) todavía no se saben derivar del SSI: esas
  celdas **no se tocan**, y solo se avisa de los **recién insertados** (los demás no interesan).
  **OJO:** marcar "PC" saca a ese CUI de las corridas siguientes (la Etapa PC lo excluye de la
  lista). Cada marcado va en el correo.
- **PROCOMPITE en una fila que ya estaba**: no se le escribe la Etapa (manda la restricción), pero
  el dato no se pierde: va al correo en una sección aparte, "el SSI los reporta como PROCOMPITE y su
  Etapa dice otra cosa", para marcarlos a mano. Mientras no se marquen, se siguen consultando y
  siguen apareciendo en ese aviso.
- **Recién insertado sin ficha del SSI**: se queda sin UEI ni Etapa **y se avisa**. Un CUI es
  "recién insertado" un solo día (el guion rearma A3:A en cada corrida), así que si ese día falla el
  SSI esas celdas ya no se llenan solas: hay que completarlas a mano o quedan vacías. Es un aviso
  distinto del de "el estado del SSI no se traduce a ninguna etiqueta": aquel es "respondió y no se
  sabe traducir", este es "no respondió".
- **Solo se escriben las celdas que CAMBIAN**, agrupadas en tramos de filas consecutivas. Una celda
  que no corresponde actualizar no entra en el envío: así no se pisa nada ni se congela una fórmula.
- Cualquier título de A2, D2, F2 o I2 que no coincida exactamente: **se detiene y se avisa**, no se
  escribe nada (antes la columna con el título malo se saltaba y las otras sí se escribían).
- **Justo antes de escribir** se vuelve a leer la columna A y los títulos. Si las filas se corrieron
  o una columna se movió durante la corrida (el Apps Script inserta filas en Proyección26), no se
  escribe nada y se avisa: escribir habría puesto cada monto en el CUI equivocado.
- Si el SSI se cayó entero, no se escribe nada.
- PIM-diario ya no se escribe. Sus columnas O, P y Q se vacían con `batch_clear` (`LIMPIAR_PIM_OPQ`)
  porque escribir `''` deja un texto vacío y `ESBLANCO()` devuelve FALSO (comprobado en la hoja).

### Comparación con la corrida anterior (agregado 25/11 tarde, corregido 26/11)
- Base = último `monto_ssi_*.json` de un **día anterior** a hoy. Las corridas del mismo día ya no
  sirven de base (una prueba de la mañana no debe ser la referencia de la tarde).
- Si la base no es del **día hábil previo**, se avisa por correo y se dice qué días faltan. Sábado y
  domingo no cuentan como falta (la PC está apagada), así que un lunes contra viernes no avisa.
- Si el archivo de la corrida anterior está dañado, se usa uno más viejo **y se avisa** (antes solo
  se imprimía en consola, que nadie ve en una tarea programada).
- Si no hay ninguna corrida anterior, se avisa: sin base, un 0 erróneo del SSI entraría a la hoja
  sin que nada lo detecte.
- Compara el `monto_final` de la base (o `monto_inversion` en archivos viejos) con lo descargado hoy.
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
- **Freno por ceros masivos** (`MAX_CEROS_SOSPECHOSOS = 5`): si 5 o más CUI que ayer tenían monto
  llegan hoy en 0, eso no es un cambio real sino una falla de la página del MEF. **No se escribe
  nada en la hoja** y se avisa (código 6). Uno o dos ceros siguen pasando como alerta nomás.
- Si `CARPETA_MONTOS` no está disponible, los archivos se guardan en `C:\Users\HP\Downloads` y se
  avisa. Un fallo al generar el Excel de variaciones ya **no** tumba la corrida (antes se perdían
  los montos y no se escribía en la hoja).
- Probado con corridas anteriores simuladas en `pruebas_sin_red.py` (alza 5 %, baja 15 %, caída a 0,
  0 repetido el día 2 sin alerta, recuperación el día 3, sin dato, antes 0, igual, nuevo) y en real a
  las 15:24 y 15:34 (0 cambios, sin Excel ni correo, como corresponde).

### Correo
- El correo del Apps Script se lee de **Configuración del proyecto > Propiedades de la secuencia de
  comandos**, propiedad `CORREO_USUARIO`. Antes la constante decía `PONER_AQUI_TU_CORREO`, así que
  **ningún aviso del guion de Google llegaba** (fallaban todos en silencio).
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

## 6. Decisiones vigentes (revisadas el 26/11)
Decisiones confirmadas por la usuaria en la revisión del 26/11:
1. **Se escribe en Proyección26, ubicando por el CUI de la columna A** (no en PIM-diario por posición).
   Esto quita de raíz el riesgo de desalineación: PIM-diario se rearma cada día y su columna N se
   corre cuando entra un CUI nuevo, pero Proyección26 se escribe fila por fila buscando el CUI.
2. **Solo se escriben las celdas que cambian**; lo que no corresponde actualizar no se envía.
3. **Monto null o sin equivalencia: la celda no se toca.** Nunca se pisa un dato bueno con 0 ni con vacío.
4. **Cero sospechoso = de > 0 a 0**, se mantiene el anterior y se avisa solo la primera vez. Riesgo
   conocido y aceptado: si un costo de verdad pasa a 0, la hoja conserva el valor viejo.
5. **5 ceros o más = falla del MEF**: no se escribe nada (`MAX_CEROS_SOSPECHOSOS`, código 6).
6. **Base de comparación = la última corrida de un día anterior**, y se avisa si falta algún día hábil.
7. **D (UEI) y F (Etapa) solo en las filas de los recién insertados**; el monto (I) en todas las
   filas de la lista. La restricción hace estructural la garantía de no pisar nada: el guion de
   Google deja D y F vacías en la fila que inserta, así que una fila vieja no se toca nunca.
   - Marcar Etapa "PC" saca al CUI de las corridas siguientes. Confirmado; cada marcado va al correo.
   - Un PROCOMPITE que **ya estaba** en la lista no se marca (es fila vieja), así que nunca saldría
     de las consultas por sí solo: se reporta por correo, en una sección aparte, para marcarlo a
     mano. Mientras no se marque, sigue apareciendo en ese aviso en cada corrida.
   - Un recién insertado para el que el SSI no respondió se queda sin UEI ni Etapa y **se avisa**:
     es su único día en la lista, así que esas celdas ya no se llenan solas.
8. **Las etapas que no sean PROCOMPITE ni desactivado no se tocan**, y solo se avisan si el CUI es
   un recién insertado. Derivar ET, EF, L... del SSI queda pendiente.
9. **UEI vacía en un recién insertado no se avisa**: es común (pedido de la usuaria).
10. **Los "NO REGISTRADO EN EL INVIERTE" no avisan** (monto 0 legítimo, hoy 6).
11. **Un solo correo por corrida**; si el correo falla, el aviso queda en un .txt.

Puntos técnicos:
- Lectura de la hoja con `UNFORMATTED_VALUE` y validación de CUI (`como_cui`: int, float entero o texto de
  7 dígitos). Ojo: `UNFORMATTED_VALUE` devuelve el **resultado** de una fórmula, no su texto.
- Escritura con `valueInputOption RAW` en tramos de filas consecutivas, en una sola llamada.
- Vaciado real de celdas con `batch_clear`: escribir `''` deja un texto vacío y `ESBLANCO()` da FALSO.
- `consultar_ssi`: clasifica respuesta (lista vacía = no existe; error de red, HTTP ≠ 200 o no-JSON = falla),
  3 intentos con pausa. Depende de `curl.exe` (solo Windows).
- Bloqueo de corridas simultáneas con `nata_apuntes\corrida_en_curso.lock` (se reusa si tiene más de
  `BLOQUEO_MAX_MIN` minutos, para que una corrida interrumpida no bloquee para siempre).
- La contraseña del correo se lee de la variable `GMAIL_APP_PASSWORD` o del registro de Windows.
- `num()` devuelve 0.0 si el importe llega como texto con separador de miles (`'1,234.56'`). Hoy el SSI
  manda números; si algún día cambia, el freno por ceros masivos lo atajaría.
- `pruebas_sin_red.py` no cubre la consulta real al SSI, la lectura real del libro ni el envío del correo.

## 7. Pendientes
1. **Tarea diaria a las 6:45** (Programador de tareas de Windows). Debe correr después de la
   descarga de Apps Script (6:20) y solo corre si la PC está prendida. Usar la ruta completa de
   python.exe y "Iniciar en" = carpeta `4-script`. Considerar "ejecutar lo antes posible si se perdió".
2. **Confirmar el libro**: `ID_LIBRO` apunta a "Copia de Drives exclusivos 19.11". Si el definitivo es
   otro, cambiar la constante y compartir ese libro con la cuenta de servicio (correo en
   `credenciales\correo iam.txt`).
3. **Verificar en el libro real** que Proyección26 tenga `UEI` en **D2** y `(A) Monto de inversión`
   en **I2** (los títulos se comparan exactos, sin tildes ni mayúsculas). Si no coinciden, el
   programa se detiene y avisa sin escribir nada.
4. **Poner `CORREO_USUARIO`** en las propiedades del proyecto de Apps Script (sin eso, los avisos
   del guion de Google solo quedan en el registro de ejecuciones).
5. **Seguridad**: la contraseña de aplicación de Gmail está en texto plano en
   `nata_apuntes\apuntes ssi monto script.txt` y `nata_apuntes\mail_notif_daily_script.md` (OneDrive) y
   pasó por el chat. Recomendado: borrarla de esos archivos, generar una nueva y volver a `setx`.
6. Ideas en `prompt_xall.md`, no pedidas aún (la comparación por CUI con la corrida anterior ya está
   hecha): revisar devengado acumulado exorbitante, ver qué fórmulas dependen de
   las columnas que se modifican, "si no está la columna Etapa, buscarla en la fila" (hoy se detiene).
7. Corregir o quitar de Proyección26 los CUI 3000882 y 3999999 (no existen en el SSI).
8. Completar a mano las UEI que el SSI trae con un nombre que no está en `UEI_EQUIVALENCIAS`
   (el correo las lista con el texto exacto que llegó, para poder agregar la equivalencia).
9. `4-script\montos_ssi.gs` ya no sirve (bloqueo); se puede borrar del proyecto de Apps Script.

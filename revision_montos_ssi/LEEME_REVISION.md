# Paquete para revisión – programa de montos del SSI (25/11/2026)

> **Estado al 26/11/2026:** la revisión se hizo y los cambios acordados ya están aplicados en
> `1-codigo/`. El programa ahora escribe en **Proyección26** (columnas D, F e I), no en PIM-diario.
> Ver la sección 6 del HANDOFF para las decisiones vigentes. Lo de abajo es el pedido original.

## Qué se pide revisar
El programa `1-codigo/cuis_desde_sheets.py`, que cada día:
1. lee dos listas de CUI de un libro de Google (hojas Proyección26 y PIM-diario) con validaciones;
2. consulta el SSI del MEF por cada CUI (monto de inversión, UEI y tipo);
3. compara con la corrida anterior (Excel de variaciones, alerta > 10 %, ceros sospechosos);
4. escribe las columnas D (UEI), F (Etapa) e I (monto) de Proyección26, ubicando cada fila por su
   CUI, y manda un solo correo con los avisos.

Buscar: errores de lógica, casos borde no cubiertos, riesgos de pisar datos buenos en la hoja, y si las
decisiones de la sección 6 del HANDOFF son razonables. Devolver hallazgos concretos (línea, caso que falla,
consecuencia) antes que sugerencias de estilo.

## Orden de lectura
1. `HANDOFF_montos_ssi.md` – contexto completo; la sección 6 lista decisiones y riesgos a mirar.
2. `2-requisitos/prompt1.md`, `prompt2.md`, `prompt3.md` – lo que pidió la usuaria, en ese orden
   (texto informal, tal cual lo escribió). Además, dos pedidos que se dieron en el chat y no están en archivo:
   - comparar con la corrida anterior: Excel con CUI, nombre, monto anterior, monto actual y % de variación;
     el correo alerta solo variaciones mayores al 10 % (constante global); si llegan ceros donde antes había
     monto, quedarse con el valor anterior y avisar;
   - un cero que se repite en corridas siguientes NO debe volver a alertar.
   `prompt_xall.md` son ideas sueltas aún no pedidas (fuera del alcance).
3. `1-codigo/cuis_desde_sheets.py` – el programa.
4. `1-codigo/pruebas_sin_red.py` – pruebas que simulan hoja, corridas anteriores y correo.
5. `1-codigo/apps_script_vigente_descarga_diaria.gs` – solo contexto: el guion de Google que llena
   PIM-diario (columnas C a L, A1, A3:A y N2:N) antes de que corra el programa.
6. `3-salidas-ejemplo/` – salidas reales de la corrida de 15:34 (lista de CUI, montos en JSON y CSV) y un
   Excel de variaciones con **datos simulados** (las corridas reales de hoy no tuvieron cambios).

## Cómo correr las pruebas (sin red ni credenciales)
```
cd 1-codigo
python pruebas_sin_red.py        # termina con "TODAS LAS PRUEBAS PASARON"
```
Requiere Python 3 con `gspread` y `openpyxl` instalados (el módulo importa gspread aunque las pruebas no
se conectan). La corrida completa del programa NO se puede hacer fuera de la PC de la usuaria: necesita la
credencial de Google, la contraseña del correo y `curl.exe` de Windows.

## Rutas originales (las copias de esta carpeta son del 25/11 15:40; si difieren, mandan las originales)
| Copia | Original |
|---|---|
| HANDOFF_montos_ssi.md | C:\Users\HP\OneDrive\daily-new-inv\HANDOFF_montos_ssi.md |
| 1-codigo/cuis_desde_sheets.py | C:\Users\HP\OneDrive\daily-new-inv\4-script\cuis_desde_sheets.py |
| 1-codigo/pruebas_sin_red.py | C:\Users\HP\OneDrive\daily-new-inv\4-script\pruebas_sin_red.py |
| 1-codigo/apps_script_vigente_descarga_diaria.gs | C:\Users\HP\OneDrive\daily-new-inv\descarga_diaria_mef_23_11_26_3pm.gs |
| 2-requisitos/*.md | C:\Users\HP\OneDrive\daily-new-inv\nata_apuntes\ |
| 3-salidas-ejemplo/monto_ssi_* | C:\Users\HP\Documents\drives_exclusivos_descarga\monto_ssi\ |
| 3-salidas-ejemplo/cuis_* | C:\Users\HP\OneDrive\daily-new-inv\nata_apuntes\test_responses\ |

## Lo que NO se incluye, a propósito
- `credenciales\credenciales.json`: clave privada de la cuenta de servicio de Google.
- `nata_apuntes\apuntes ssi monto script.txt` y `nata_apuntes\mail_notif_daily_script.md`: contienen la
  contraseña de aplicación de Gmail en texto plano.
- Ojo: el programa sí incluye el correo de aviso de la usuaria y el ID del libro de Google (constantes al
  inicio). El libro no es público; el ID solo no da acceso.

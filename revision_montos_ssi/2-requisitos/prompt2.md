considerar: 

- evitar trabajar en hoja de google para no ralentizarla para otros usuarios.

modifica script para que:

- descargar datos en base a CUIs todos_cuis y recien_insertados:
  
  - recien_insertados estan incluidos en todos_cuis.
    - - para recien_insertados: traer dato de: 1.UEI 2.tipo inversion 3. monto inversion. 
      - para todos_cuis: solo traer monto inversion.

- casos:
  
  - si al buscar cui no encuentras pagina de ese cui o su servicio entonces reportar por mail.
  - si monto inversion viene vacio le pones 0. 

- en "C:\Users\HP\Documents\drives_exclusivos_descarga\monto_ssi" poner valores descargados diarios (guarda dato de fecha y hora por si ejecuto varias veces al dia cuando hago pruebas).

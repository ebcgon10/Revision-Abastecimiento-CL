# Sistema de Revisión Abastecimiento — CD Coquimbo

App Streamlit, del mismo estilo que *Sistema de Revisión Venta*, que prioriza el masivo
CL → CD (COQUIMEXT → COQUIMBO) con la **venta real del CD** en vez de solo el forecast del
reporte stock 7 hrs.

## Qué se sube cada día

1. **Modelo de abastecimiento** (`ABASTECIMIENTO_x.xlsx`): el mismo Excel que genera el masivo.
   La app lee las hojas ABASTECIMIENTO, Stock a Mantener, REVISION, DICCIONARIO,
   MASIVO GM, MASIVO CVP y los masivos de tránsito.
2. **Venta real**: exporte SAP 951 (`Fecha contabiliz.`, `Material`, `Cantidad`;
   cantidad negativa = salida) en CSV o Excel. Se pueden subir varios meses; las filas repetidas
   entre archivos se cuentan una sola vez. También acepta un formato simple `FECHA`, `CODIGO`, `CAJAS`.
   Si la app tiene los secrets de Drive, toma la venta directo de la carpeta (archivos con
   "VENTA" en el nombre).

## Cómo calcula

- **Demanda/día** = cajas vendidas en las últimas N semanas ÷ días operativos (sin domingo).
  Con *alza de la última semana* activado usa la mayor entre ese promedio y el de la última semana.
- **Cobertura CD** = stock CD ÷ demanda/día.
- **Pallets sugeridos** = los justos para llegar a los días objetivo (tope: pallets del masivo).
- **Prioridad de cada línea del masivo**:
  - `1-URGENTE`: cobertura menor al umbral (2 días por defecto).
  - `2-NORMAL`: le falta para llegar a los días objetivo.
  - `3-POSTERGAR`: ya tiene los días objetivo; no sube hoy.
  - `4-NO TRAER`: sin venta en el período.
- **Cruce completo** (todos los SKU):
  - `AGREGAR AL MASIVO`: urgente, hay pallet en CL y el modelo no lo pidió.
  - `SOBRA EN CD`: cobertura sobre 15 días; muestra cuántos pallets están de más (candidatos a
    devolver o a no reponer para liberar strech).
  - `QUIEBRE (sin pallet en CL)`, `BAJO (viene en masivo)`, `VENDE Y NO ESTÁ EN EL MODELO`, `OK`.
- **Sobra en CD → enviar a CL**: productos sobre el umbral de exceso, de mayor a menor pallets
  sobrantes (deja en el CD los días objetivo y el resto se envía). Descarga aparte con la hoja
  `TRASLADO CD - CL` en el mismo formato de la hoja CD - CL del modelo.
- **Forecast vs venta real**: los SKU donde el forecast más se aleja de lo vendido, en pallets
  a 5 días, para corregir la hoja REVISION.

El Excel descargado trae `MASIVO PRIORIZADO` con el mismo formato del masivo (GRUPO, ORIGEN,
DESTINO, FECHA, CODIGO SAP, CANTIDAD CAJAS, PALLETS, PASILLO ACTUAL) solo con los pallets
sugeridos y en orden de prioridad, más `DETALLE MASIVO`, `CRUCE COMPLETO` y `RESUMEN`.

## Publicar en Streamlit Cloud

1. Sube esta carpeta a un repositorio de GitHub (por ejemplo `ebcgon10/Revision-Abastecimiento`).
2. En share.streamlit.io crea la app con `app.py` como archivo principal.
3. (Opcional, para leer la venta desde Drive) en *Settings → Secrets* pega:

```toml
gcp_json = '''{ ...contenido del JSON de la cuenta de servicio de la torre de control... }'''
drive_venta_folder_id = "18xG-E755XZcz9f-eCd1DpboJg8gKwxsN"
```

   y comparte esa carpeta de Drive con el correo de la cuenta de servicio (lector basta).
   Sin secrets, la app funciona igual subiendo la venta a mano.

## Correr en el computador

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Archivos

- `app.py` — interfaz.
- `motor.py` — toda la lógica (sin Streamlit; se puede probar o automatizar aparte).
- `drive.py` — lectura de Drive con cuenta de servicio (el mismo de la torre de control).
- `tema.py` — tema CCU verde/blanco común a las apps del CD.

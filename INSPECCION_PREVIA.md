# Inspeccion previa (2026-03-13)

Fuente usada:
- Export XLSX de la hoja `1MynyGPeRhpuo57XbjskdwVxoqgkmpfVwsPE4i7eRycM`
- Reporte completo generado en: `C:\Users\elrub\Desktop\CARPETA CODEX\reports\sheet_inspection_report.json`

## 1) Estructura completa (libro, pestanas, rangos usados, filas/columnas)
Libro: `(NO USAR - APP EN DESARROLLO) CRM: VENTA-BOOKING`

Pestanas (23):
- INFANTIL: `A1:L2316` (2316x12)
- IGUALDAD: `A1:L3142` (3142x12)
- CTRO.CULTURA: `A1:O3773` (3773x15)
- TEATRO: `A1:AB4222` (4222x28)
- ABUELOS: `A1:M4821` (4821x13)
- CHIRINGUITOS: `A1:M1103` (1103x13)
- CAMPING: `A1:N1818` (1818x14)
- Club & casinos: `A1:N1818` (1818x14)
- FESTEJOS: `A1:L3112` (3112x12)
- CULTURA: `A1:L8004` (8004x12)
- HOTELES: `A1:P7983` (7983x16)
- JUVENTUD: `A1:L9021` (9021x12)
- BODAS: `A1:K1955` (1955x11)
- CLUB: `A1:N1878` (1878x14)
- AYTO_UD-CULT: `A1:K109` (109x11)
- GDOR.PLIEGO: `A1:M6027` (6027x13)
- EMP PROGRAMACION: `A1:M4073` (4073x13)
- EVENTOS EMP TOP: `A1:M4050` (4050x13)
- PROMO_EMP-PROG: `A1:I1023` (1023x9)
- EVENTOS EMPRESAS: `A1:K3212` (3212x11)
- CCAA: `A1:B51` (51x2)
- GALAS DE PREMIOS: `A1:J1025` (1025x10)
- ROOFTOPS: `A1:K352` (352x11)

## 2) Datos visibles relevantes y formulas
- Se detectaron encabezados y datos visibles en todas las hojas.
- Ejemplo de encabezados recurrentes: `CCAA`, `PROVINCIA`, `MUNICIPIO`, `EMAIL`, `TELEFONO`, `CORREO ENVIADO`, `LLAMADA`, `Merge status`.
- Formulas detectadas en: `CAMPING` (8), `Club & casinos` (3), `HOTELES` (15), `CLUB` (1).

## 3) Formatos aplicados
- Tipografias dominantes: `Arial` y `Google Sans Text`.
- Formato de numero dominante: `General`, con uso puntual de `#,##0`, `#,##0"EUR"` y formatos de fecha tipo `d" de "mmmm`.
- Alineacion dominante: sin alineacion explicita (default).
- Relleno dominante: sin color explicito (default).
- Bordes: no se detectaron celdas con borde explicito en el export.
- Anchos de columna personalizados: presentes en multiples hojas.

## 4) Celdas combinadas
- Celdas combinadas detectadas: `0` en las 23 hojas.

## 5) Validaciones de datos y desplegables
- Validaciones detectadas en hojas clave (ej.: INFANTIL, IGUALDAD, TEATRO, CAMPING, CLUB, EVENTOS EMPRESAS).
- Regla tipica de desplegable detectada: `"PTE,SI,NO COGE,NO EXISTE"`.
- Hojas sin validaciones detectadas en el export: por ejemplo `HOTELES`, `GDOR.PLIEGO`, `EMP PROGRAMACION`, `EVENTOS EMP TOP`, `PROMO_EMP-PROG`, `CCAA`, `GALAS DE PREMIOS`, `ROOFTOPS`.

## 6) Filtros y vistas de filtro
- Auto filtros detectados en multiples hojas (rango de filtro presente en metadata XLSX).
- Ordenaciones (sort state) detectadas en varias hojas.
- Vistas de filtro de Google Sheets: no se exportan de forma fiable en XLSX.

## 7) Reglas de formato condicional
- Regla general: `8` reglas por hoja en casi todo el libro (excepto `CCAA`, con `0`).
- Tipo de regla detectada: `cellIs`.
- Ejemplos de valores objetivo en reglas: `EMAIL_OPENED`, `EMAIL_CLICKED`, `RESPONDED`, `BOUNCED`, `ERROR`, `NO_RECIPIENT`.

## 8) Protecciones de hoja/rango y permisos
- Proteccion de hoja detectada en XLSX: no activa en las hojas exportadas.
- Seguridad de libro detectada en XLSX: `lockStructure=false`, `lockWindows=false`.
- Permisos de editores/lectores de Google Sheets: no se exponen en XLSX.
- Mecanismo tecnico preparado para extraer permisos/protecciones reales desde Apps Script:
  - Archivo: `INSPECCION_HOJA.gs`
  - Funcion: `codexInspeccionarHojaVinculada`
  - Incluye: permisos de archivo (owner/editors/viewers), protecciones de hoja/rango, validaciones, condicionales y detalle API avanzada (si esta habilitada).

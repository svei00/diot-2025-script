# HANDOFF — diot-2025-script

Documento autosuficiente para retomar el proyecto en una sesión nueva
(Sonnet u Opus) sin el historial de la conversación. Última actualización:
**2026-11-11**.

## 1. Qué es y de dónde salen las reglas

`diot_generator.py` + paquete `diot/` leen el libro **Registro de
Facturación 2026v5** (`.xlsb/.xlsm/.xlsx`, solo lectura) y generan el TXT de
carga masiva de la DIOT 2025+: un renglón por proveedor (tipo de tercero 04),
54 campos separados por `|`, UTF-8, CRLF.

- **Libro vigente** (solo lectura, nunca modificarlo):
  `D:\repos\registro-facturacion-xlm\doctos\Registro de Facturación 2026v5.xlsm`
- **Fuente de verdad de la metodología:** `HANDOFF.md` §15 del repo
  `registro-facturacion-xlm` (rama `claude/formulas-validacion-sat-3da366`
  hasta que se mezcle). Lo indispensable está copiado aquí, en §3.
- **Oráculo de regresión** (sin versionar, `resources/`): copia del libro del
  12-ago-2026 y el TXT de julio 2026 que se presentó.
- **Fuera de alcance:** cfdi-app, papeles-trabajo, importar_xml_a_contabilidad,
  el libro y el repo registro-facturacion-xlm.

## 2. Arquitectura

```
diot_generator.py        entrada: CLI y ventana; reexporta build_diot y read_sheet
diot/valores_celda.py    texto, número, "Sí", fechas dd/mm/aaaa (sin locale)
diot/redondeo.py         Art. 20 CFF y base mínima del validador SAT (16% y 8%)
diot/lectura_libro.py    lectura .xlsb/.xlsm/.xlsx (valores en caché), hoja Control
diot/columnas.py         columnas por encabezado; opcionales solo por nombre exacto
diot/facturas.py         universo de facturas, bucket A/B, importes por tasa, exentos
diot/rep.py              universo de REP, EquivalenciaDR, importes, REP sobre PUE
diot/region_fronteriza.py  aviso de CP de la frontera sur
diot/acumulados.py       ImportesPorTasa y TotalesDeBucket
diot/layout_txt.py       números de campo oficiales y armado del renglón
diot/construir.py        build_diot(): PURA (sin I/O, sin tkinter) — Boardflare
diot/resumen_libro.py    Resumen por encabezado (bloque fila 1 + subtítulo fila 2)
diot/declaracion_libro.py  Declaracion por etiqueta (columna B), solo mismo mes
diot/conciliacion.py     texto de conciliación por tasa y avisos
diot/generar.py          orquestación con I/O (leer, build, escribir, conciliar)
diot/archivo_salida.py   nombre del TXT, config por RFC, diot_rfc_op.csv
diot/ventana.py          tkinter
tests/                   pytest con filas sintéticas (nunca el libro real)
DIOT_Macro.bas           port VBA de la misma lógica
```

`build_diot(rec_headers, rec_rows, pag_headers, pag_rows, year, month, op_map,
limite_efe=2000, usos_ded=None)` regresa `(renglones, ResultadoDiot)`.
`ResultadoDiot` trae `buckets` (A, B, C con `numero_documentos` e
`importes`), ajustes de base por tasa, `rep_descartados` y `avisos`.

## 3. Reglas (equivalen a §4.2 del prompt original)

**Universo de facturas** (`RecibidasXML`, buckets A y B): RFC presente;
Fecha Emision en año/mes; `Tipo = "Factura"`; `Estado ≠ "Cancelado"` (vacío =
vigente); método PUE; UsoCFDI ∈ usos de Control (G01, G03); **`Total > 0`**
(el prellenado excluye remanentes de anticipo y facturas de $0.01 con 100% de
descuento; cambia el conteo, no el IVA).

- **A** bancarizada: `Bancarizado = "Sí"`.
- **B** efectivo: `FormaDePago 01`, `Total ≤ LimiteEfectivo` (IVA incluido),
  combustible siempre fuera.

**Universo de REP** (`PagosRecibidasXML`, bucket C, sin cambio): FechaPago en
año/mes, `Estado ≠ "Cancelado"`, `EsDuplicado ≠ 1`, `RelUUID = "OK"`,
`BancarizadoP = "Sí"`. `RelUUID ≠ OK` se excluye y se reporta.

**Moneda:** si `MonedaDR ≠ MXN`, `ImpPagado` y todos los importes de
impuestos del REP vienen en la moneda del documento → se dividen entre
`EquivalenciaDR` (vacía o 0 → 1). Caso real octubre: BUCA RECUBRIMIENTOS,
USD, equivalencia 0.0544401406, base 2,045.50 → 37,573.38. Esa era toda la
DIF −5,684.46 del bucket C. Si `MonedaP ≠ MXN`: se excluye y se reporta
(la conversión necesitaría `TipoCambioP`; hoy no hay casos).

**Tasa 16%:** base = IVA 16% ÷ 0.16 en facturas (A y B); REP `IVA 16 Base`.

**Tasa 8% (estímulo región fronteriza, desde octubre 2026):**
- A: base = SubTotal − Descuento (cuadra al centavo con el prellenado).
- B: base = IVA 8% ÷ 0.08 (no el neto: puede traer IEPS o exentos).
- Factura con IVA 16% e IVA 8% a la vez: base 8% = IVA ÷ 0.08 y **aviso**.
- REP: `IVA 8 Base` / `IVA 8 Importe` ÷ EquivalenciaDR.

**Exentos (campo 50):** facturas del bucket A sin IVA 16% ni 8%, con
`Total Trasladados = 0`, que son colegiatura (`Nivel Educativo` lleno) o cuyo
`Conceptos` contiene "Alumno": base = SubTotal − Descuento (heurística de
`Resumen!AT`). REP: `IVA Excento Base` ÷ EquivalenciaDR. Las facturas sin
traslados que no son colegiatura (cuotas IMSS, "no objeto") **no** son
exentas ni tasa 0%.

**Tasa 0% (campo 51):** solo REP `IVA Cero Base` ÷ EquivalenciaDR. La de
facturas (octubre SAT: 8,991) no se puede derivar del export: son conceptos a
0% dentro de facturas mixtas (aerolíneas, gasolineras).

**IVA retenido (campo 48):** facturas `Retenido IVA` + REP `Ret IVA * Importe`
÷ EquivalenciaDR, con aviso.

**Redondeo:** por proveedor, Art. 20 CFF (0.01–0.50 baja, 0.51–0.99 sube).
Luego la base de cada tasa sube al mínimo entero con
`(base × tasa%) // 100 ≥ IVA` (el validador del SAT trunca).

**Fechas:** texto dd/mm/aaaa; `anio_y_mes()` las parsea sin locale. Nunca
DATEVALUE ni nada que dependa de la configuración regional.

**Lectura de hojas para conciliar:** `Resumen` por encabezado (filas 1 y 2;
falla con mensaje claro si falta una columna obligatoria; `Resumen!Y` ya no es
base de IVA, es "Pagado s/IVA"). `Declaracion` por etiqueta y solo si su mes
(C2/D2) es el que se genera.

## 4. Layout de campos (P1) — confirmado

Fuente: *Instructivo para el armado del archivo de carga masiva – DIOT*, SAT,
enero 2025, sección 3, transcrito en `D:\repos\importar_xml_a_contabilidad\diot.py`
(encabezado). Svei confirmó el mapa el 2026-11-11.

| Campo | Contenido |
|---|---|
| 1 / 2 / 3 | Tipo de tercero (04) / tipo de operación (85, 03, 06) / RFC |
| 8 / 9 | Valor actos pagados / DDB — región fronteriza **norte** 8% |
| 10 / 11 | Valor actos pagados / DDB — región fronteriza **sur** 8% |
| 12 / 13 | Valor actos pagados / DDB — tasa 16% |
| 14–17 | Importaciones (aduana tangibles; intangibles y servicios) |
| 18 / 20 / 22 | IVA acreditable exclusivamente gravadas — norte 8% / sur 8% / 16% |
| 19 / 21 / 23 | IVA acreditable asociado a proporción — norte / sur / 16% |
| 28–47 | IVA no acreditable (por tasa y causa) |
| 48 | IVA retenido por el contribuyente |
| 49 | Actos pagados en importación exentos |
| 50 | Actos pagados exentos |
| 51 | Demás actos pagados a tasa 0% |
| 52 / 53 | No objeto (territorio nacional / sin establecimiento) |
| 54 | Manifiesto de efectos fiscales (no se llena; el TXT de julio se aceptó así) |

El generador llena 1, 2, 3, 8, 12, 18, 22, 48, 50, 51.

## 5. Decisiones (2026-11-11, Svei)

| # | Pregunta | Decisión |
|---|---|---|
| P1 | Números de campo | Mapa oficial de §4 |
| P2 | Región del 8% | Fija: **norte** (campos 8 y 18). Octubre: todos los CP de expedición son 21xxx (Mexicali) o 84xxx (SLRC). Aviso si un CP empieza con 24, 29, 30, 77 u 86 (estados de la frontera sur) |
| P3 | Exentos | Se reportan ya, por proveedor, con la heurística del libro. Cambiar a base por tasa del XML cuando exista la Adenda 2 |
| P4 | Tasa 0% de facturas | 0 en el TXT, con nota y cruce "capturar a mano" en la conciliación |
| P5 | REP sobre factura PUE | Igual que el SAT (cuenta dos veces) + aviso. El método se busca en la factura relacionada por UUID (Pagos 2.0 no trae `MetodoDePagoDR`) |
| P6 | Retenciones IVA en REP | Sí, al campo 48, convertidas y con aviso |
| P7 | DIOT_Macro.bas | Se actualiza a la par (Svei prefirió no deprecarla) |
| Git | Ramas | Todo directo en `main`, sin ramas; commits sin coautoría |

## 6. Pruebas

Automáticas: `python -m pytest` (41 pruebas, filas sintéticas: USD, 8%,
exentos, Total $0, efectivo con IEPS, REP excluidos, factura mixta 16%+8%,
REP sobre PUE, retenciones de REP, Resumen/Declaracion por encabezado y
etiqueta, conciliación).

Manuales contra el libro real (no se automatizan: tiene datos de clientes):

```bash
python diot_generator.py 7 N "jul.txt" --libro "resources\Registro de Facturación 2026v5.xlsm"
fc /b jul.txt "resources\07. Jul 2026  N DIOT Declaración.txt"
python diot_generator.py 10 N "oct.txt" --libro "D:\repos\registro-facturacion-xlm\doctos\Registro de Facturación 2026v5.xlsm"
```

| Prueba (2026-11-11) | Esperado | Resultado |
|---|---|---|
| Julio 2026 (copia 12-ago) vs TXT presentado | byte-idéntico | ✓ en fases 1–4 |
| Oct — A documentos | 405 | 405 ✓ |
| Oct — C IVA 16% (convertido) | 29,678.97 | 29,678.97 ✓ |
| Oct — base/IVA 8% A | 192,641.81 / 15,411.31 | idéntico ✓ |
| Oct — base/IVA 8% B | 46.25 / 3.70 | idéntico ✓ |
| Oct — base/IVA 8% C | 172,449.96 / 13,796.00 | idéntico ✓ |
| Oct — exentos facturas / REP | 14,440 / 520 | idéntico ✓ |
| Oct — total base/IVA 16% | 4,507,075.69 / 721,132.12 | 4,507,075.69 / 721,132.11 ✓ (centavo de redondeo) |
| Oct — total base/IVA 8% | 365,138.02 / 29,211.01 | idéntico ✓ |
| Declaracion "Monto a capturar 16%" | 4,507,075.69 | OK |
| Declaracion "Monto a capturar 8%" | 365,138.02 | OK |
| Declaracion "TOTAL IVA ACREDITABLE" | 750,343.12 | 750,343.12 OK |
| Ene–sep (libro vigente) fase 2 vs fase 1 | sin cambio | ✓ |
| Fase 3 vs fase 2 | solo exentos | mar, ago, sep, oct: +1 renglón con campo 50 (= Resumen!AT/AU) |
| TXT octubre, suma por campo | — | 12: 4,507,233 · 22: 721,134 · 8: 365,176 · 18: 29,212 · 50: 14,960 |

Ajustes de base por la invariante en octubre: +163 pesos (16%) y +38 (8%).

## 7. Pendientes

1. **Macro VBA sin probar en Excel.** Se portó desde el Python, pero aquí no se
   puede ejecutar VBA. Correr octubre 2026 con la macro y comparar el TXT byte
   a byte contra el de Python antes de usarla.
2. **Tasa 0% de facturas** y **exentos exactos**: dependen de la Adenda 2
   (columnas `Base IVA 16%/8%/0%/Exento` leídas del nodo `Traslado` del XML).
   Cuando existan, cambiar `facturas.importes_tasa_8`, `importe_exento` e
   incluir la tasa 0% en `importes_de_factura`, y quitar la heurística.
3. **Meses ya presentados (marzo, agosto, septiembre):** con la regla de
   exentos el TXT agrega el campo 50 en un renglón. Si se presentan
   complementarias, decidir si se incluyen.
4. **`MonedaP ≠ MXN`**: hoy se excluye y se reporta. Si aparece un caso real,
   definir la conversión con `TipoCambioP`.
5. **Notas de crédito recibidas** (campo 13, DDB 16%): no se llenan; el libro
   tampoco las separa por proveedor para la DIOT.
6. **Región sur**: si un proveedor de la frontera sur aparece, decidir si se
   pasa a regla por CP (campos 10 y 20).

## 8. Mensajes de commit de esta entrega

```
814f732 fix(diot): fase 1 - correcciones sin cambio de layout
cc882fa feat(diot): fase 2 - IVA 8% región fronteriza norte (campos 8 y 18)
4c44706 feat(diot): fase 3 - exentos, tasa 0% de REP y retenciones de REP
2f0786e feat(diot): fase 4 - conciliación por tasa contra Resumen y Declaracion
        docs/feat: fase 5 - macro VBA, README y HANDOFF
```

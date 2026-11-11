# DIOT Generator — Registro de Facturación 2026v5

Generador del archivo **TXT de carga masiva de la DIOT 2025+** (Declaración Informativa de Operaciones con Terceros) a partir del `Registro de Facturación 2026v5` en **`.xlsb`, `.xlsm` o `.xlsx`**. Lee el libro en **modo lectura** (nunca lo modifica) y escribe solo el `.txt` de 54 campos que pide el SAT.

La metodología replica el prellenado del SAT (DyP) tal como quedó en el libro el 11-nov-2026. Los detalles, decisiones y pruebas están en [HANDOFF.md](HANDOFF.md).

## Características

- ✅ **Flujo de efectivo** (LIVA 1-B y 5): PUE bancarizadas, efectivo ≤ $2,000 y REP (PPD pagados en el mes)
- ✅ **Tasa 16%, región fronteriza norte 8%, exentos, tasa 0% de REP e IVA retenido**
- ✅ **REP en moneda extranjera** convertidos con `EquivalenciaDR`
- ✅ **Redondeo Art. 20 CFF** y **base mínima que exige el validador del SAT** (`trunc(base × tasa) ≥ IVA`, para 16% y 8%)
- ✅ **Conciliación** por tasa contra `Resumen` (por encabezado) y contra `Declaracion` (por etiqueta)
- ✅ **Avisos**: REP descartados, REP sobre factura PUE, facturas que mezclan 16% y 8%, CP de frontera sur, retenciones en REP
- ✅ Ventana (tkinter) o línea de comandos; nombre de archivo `MM. Abr AAAA  N DIOT Declaración.txt`

## Requisitos

- **Python 3.7+**
- `pip install pyxlsb openpyxl` (`pyxlsb` solo para `.xlsb`; `openpyxl` para `.xlsm`/`.xlsx`)
- Para las pruebas: `pip install pytest`

> **`.xlsm` / `.xlsx`:** se leen los **valores en caché** de las fórmulas (lo que Excel escribe al guardar). Si el archivo nunca pasó por Excel, el script lo dice: ábrelo, guárdalo y vuelve a correrlo.

## Uso

### Ventana (recomendado)
```bash
python diot_generator.py
```
También con doble clic en `Generar_DIOT.bat`. Eliges el libro, el mes y el tipo (`N` normal, `C1`, `C2`… complementaria), y al terminar ves la conciliación.

### Línea de comandos
```bash
python diot_generator.py 10 N "C:\out\oct.txt" --libro "D:\ruta\Registro de Facturación 2026v5.xlsm"
```
```
python diot_generator.py [mes 1-12] [N|C1|C2|...] [ruta_salida.txt] [--libro ruta]
(sin argumentos abre la ventana; sin ruta de salida abre "Guardar como";
 sin --libro busca uno en la carpeta actual: .xlsb > .xlsm > .xlsx)
```

## Parámetros que se leen de la hoja `Control`

| Parámetro | Etiqueta en `Control` | Valor 2026 |
|---|---|---|
| Año fiscal | `Año fiscal` | 2026 |
| Límite de pago en efectivo (**IVA incluido**) | `Límite pago en efectivo` | 2,000 |
| Usos CFDI deducibles | `Usos CFDI deducibles` | G01, G03 |

## Campos que se llenan

Fuente: *Instructivo para el armado del archivo de carga masiva – DIOT*, SAT, enero 2025, sección 3. Un renglón por proveedor (tipo de tercero 04), 54 campos separados por `|`, UTF-8, CRLF.

| Campo | Contenido | Origen |
|---|---|---|
| 1 | Tipo de tercero `04` (nacional) | fijo |
| 2 | Tipo de operación (`85` por omisión; `03`/`06` vía `diot_rfc_op.csv`) | catálogo |
| 3 | RFC del proveedor | factura / REP |
| 8 | Valor de actos pagados — región fronteriza **norte** 8% | facturas + REP |
| 12 | Valor de actos pagados — tasa 16% | facturas + REP |
| 18 | IVA acreditable — región fronteriza norte 8% | facturas + REP |
| 22 | IVA acreditable — tasa 16% | facturas + REP |
| 48 | IVA retenido | facturas (`Retenido IVA`) + REP (`Ret IVA * Importe`) |
| 50 | Actos exentos | facturas (heurística) + REP (`IVA Excento Base`) |
| 51 | Actos a tasa 0% | solo REP (`IVA Cero Base`) |

Los campos sin dato quedan vacíos (con su `|`). Los de 16% y 8% solo se escriben si el proveedor tiene esa tasa.

## Reglas por bucket

**Universo de facturas** (`RecibidasXML`): Fecha Emision en el mes, `Tipo = Factura`, `Estado ≠ "Cancelado"`, método PUE, UsoCFDI deducible, **`Total > 0`** (el prellenado excluye las de $0).

| | Bucket A — bancarizadas | Bucket B — efectivo ≤ límite | Bucket C — REP |
|---|---|---|---|
| Filtro | `Bancarizado = Sí` | `FormaDePago 01`, `Total ≤ 2,000`, sin combustible | `FechaPago` en el mes, `Estado ≠ Cancelado`, `BancarizadoP = Sí`, `EsDuplicado ≠ 1`, `RelUUID = OK` |
| Base 16% | IVA 16% ÷ 0.16 | IVA 16% ÷ 0.16 | `IVA 16 Base` ÷ EquivalenciaDR |
| Base 8% | SubTotal − Descuento | IVA 8% ÷ 0.08 | `IVA 8 Base` ÷ EquivalenciaDR |
| Exentos | colegiaturas / "Alumno" sin traslados | — | `IVA Excento Base` ÷ EquivalenciaDR |
| Tasa 0% | — (captura manual) | — | `IVA Cero Base` ÷ EquivalenciaDR |

- **Moneda extranjera:** si `MonedaDR ≠ MXN`, todos los importes del REP vienen en la moneda de la factura y se dividen entre `EquivalenciaDR` (vacía o 0 → 1). Si `MonedaP ≠ MXN`, el REP se excluye y se reporta.
- **Factura con 16% y 8% a la vez:** la base 8% se toma como IVA ÷ 0.08 y se avisa (hay que verificar contra el XML).
- **REP que paga una factura PUE:** se cuenta dos veces, igual que el SAT, y se avisa para pedir la corrección al proveedor.
- **Facturas sin impuestos trasladados** (cuotas IMSS, "no objeto") no son exentas ni tasa 0%.
- **Fechas:** texto `dd/mm/aaaa`; nunca se interpretan con la configuración regional.

## Conciliación — octubre 2026

```
 Bucket                 Concepto             Script          Resumen
 A PUE bancarizadas     IVA 16%           685197.03        685197.03   OK
 B efectivo <= 2000     IVA 16%             6256.12          6256.12   OK
 C REP (PPD pagados)    IVA 16%            29678.97         29678.97   OK
 A PUE bancarizadas     base 8%           192641.81        192641.81   OK
 C REP (PPD pagados)    base 8%           172449.96        172449.96   OK
 A PUE bancarizadas     exento             14440.00         14440.00   OK
 C REP (PPD pagados)    exento               520.00           520.00   OK
 A PUE bancarizadas     documentos              405              405   OK
 ...
 Totales A+B+C (antes de redondeo)              Base              IVA
 Tasa 16%                                 4507075.69        721132.11
 Región fronteriza norte 8%                365138.02         29211.01
 Exentos                                    14960.00
 IVA acreditable total (16% + 8%)                           750343.12
 ...
 vs Declaracion (por etiqueta)                Script      Declaracion
 Actos tasa 16% (con efectivo)            4507075.69       4507075.69   OK
 Actos frontera 8% (con efectivo)          365138.02        365138.02   OK
 Actos exentos (facturas + REP)             14960.00         14960.00   OK
 Actos tasa 0% (solo REP)                       0.00          8991.00   facturas 0%: capturar a mano (8991.00)
 TOTAL IVA acreditable                     750343.12        750343.12   OK
```

Si algún renglón sale `DIF`, el TXT no cuadra con el libro: revísalo antes de subirlo. La tasa 0% de facturas no se puede derivar del export actual y se captura a mano en el portal.

## Errores del SAT: "El IVA acreditable no puede ser mayor al IVA pagado"

El validador recalcula el IVA pagado del renglón desde la base y **trunca**. En mayo 2026 rechazó renglones con `round(base × 0.16) = IVA` pero `trunc(base × 0.16) < IVA` (p. ej. base 1254, IVA 201). El script sube la base al mínimo entero que cumple `(base × tasa) // 100 ≥ IVA`, con aritmética entera, para 16% y para 8%. Costo típico: unos pesos por renglón (octubre 2026: +163 en 16% y +38 en 8%).

## Clasificación de operaciones

Todos los proveedores salen como `85` (Otros). Para cambiarlo, crea `diot_rfc_op.csv` junto al libro (ver `diot_rfc_op.csv.example`):
```csv
RFC,Op
TME840315KT6,03
```

## Desarrollo

```
diot_generator.py        entrada (CLI / ventana)
diot/                    lógica, un archivo por responsabilidad (mapa en diot/__init__.py)
  construir.py           build_diot(): función pura, sin I/O (llamable desde Boardflare)
tests/                   pytest con filas sintéticas (no leen el libro real)
DIOT_Macro.bas           versión VBA que replica la misma lógica
```

```bash
python -m pytest
```

La prueba contra el libro real (octubre 2026 y julio 2026 byte a byte) es manual y está documentada en [HANDOFF.md](HANDOFF.md).

## Licencia

MIT. Úsalo, modifícalo, comparte. Ver [LICENSE](LICENSE).

## Autor & Créditos

Generado para GOOH841231EPA (Hugo Edgar González Orozco).
Con asistencia de Claude (Anthropic).

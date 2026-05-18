# SEGUFIJA-APP — Guía de contexto para Claude

## URLs y referencias clave
- **Apps Script (producción):** `https://script.google.com/macros/s/AKfycbxO1svm3KUuIwPnuwGJejtHVb3voCacNDaxSFUXgPvvcHZnO0caC7Wka0jN2L_CMpF7/exec`
- **Sheet ID:** `1YFu6po3AC_bg0k6lOMTy_IuBnVJX5Ng9_28C9TNWPYg`
- **GitHub Pages:** `anaferruz-jpg.github.io/almacen/`
- **Apps Script project ID:** `1AyMUiE92ZP9BbEEAr0zZnVl-650OUr6l62YR4Mb-eDNtDI9mB8OGFGhl`
- **Drive Ofertas folder ID:** `1OtYCuQM-Py8bjfTJkFjAP9DB4L2cZwZB`
- **Archivos locales:** `C:\Users\Segu Fija\Desktop\SEGUFIJA-APP\`
- **Versión Apps Script desplegada:** v81 (18 may 2026)

## Archivos del proyecto
| Archivo | Descripción |
|---|---|
| `index.html` | App almacén (GitHub Pages) — inventario, historial, pedidos |
| `facturacion.html` | App CRM/facturación (local + GitHub Pages) |
| `code_google_script.gs` | Copia local del Apps Script (referencia) |
| `CLAUDE.md` | Este archivo — contexto rápido para Claude |

## Arquitectura general
- **Frontend:** HTML/CSS/JS puro, sin frameworks, una sola página por app
- **Backend:** Google Apps Script (doGet / doPost) con Google Sheets como BD
- **Deploy:** GitHub Pages (`anaferruz-jpg.github.io/almacen/`) — push a `main` → ~1 min
- **POST:** `postAPI({action, ...datos})` → `fetch(URL, {method:'POST', body:JSON.stringify(...)})`
- **GET:** `getAPI(action)` → `fetch(URL + '?action=xxx')`

## Hojas de Google Sheets
| Hoja | Columnas clave |
|---|---|
| `Inventario` | Nombre, Stock, Necesario, PVP, PrecioCompra, StockMinimo |
| `Historial` | Tipo, Fecha, Obra, Resp, Prov, Ref, Obs, Items(JSON), Anulada, EstadoPedido, FechaLimite |
| `Obras` | Nombre, Lugar, Presupuesto, Inicio, Fin, Estado, Obs, Partidas(JSON), Personal(JSON), PctAlquiler, DiasAlquiler, PrevisionPersonal(JSON), Tipo, ML, Meses, FormaPago, Plazo, Contacto(JSON), CodObra, Constructora, CarpetaDriveId, PlanFac(JSON) |
| `CRM_Obras` | oferta, constructora, obra, estado, prioridad, importe, ofertaEnviada, contacto, telefono, email, ultimaAccion, proximaAccion, fechaRec, carpetaDriveId |
| `Facturas` | id, fecha, cliente, obra, concepto, base, pctIva, iva, total, estado, vencimiento |
| `Gastos` | id, fecha, proveedor, concepto, base, pctIva, iva, total, estado, vencimiento, obra, categoria |
| `GastosInternos` | id, fecha, concepto, base, pctIva, iva, total, importe, estado, vencimiento, obra, categoria |
| `Clientes` | nombre, cif, direccion, email, telefono |
| `Proveedores` | nombre, cif, direccion, email, telefono |
| `Operarios` | nombre, categoria, precioHora |
| `Presupuestos` | id, fecha, cliente, obra, items(JSON) |

## Apps Script — acciones registradas
### doGet
`getInventario`, `getHistorial`, `getObras`, `getFacturas`, `getGastos`, `getClientes`,
`getProveedores`, `getCobros`, `getGastosInt`, `getCRM`, `getDocsPersonal`,
`getDocsGestoria`, `getDocsObra`, `getOperarios`, `getPresupuestos`, `getCostesMateriales`

### doPost
`updateStock`, `restoreStock`, `addHistorial`, `updatePrecios`, `addArticulo`, `editArticulo`,
`deleteArticulo`, `saveObra`, `deleteObra`, `deleteHistorial`, `marcarPedidoEntregado`,
`editarHistorial`, `saveFactura`, `deleteFactura`, `saveGasto`, `deleteGasto`,
`saveGastoInterno`, `deleteGastoInterno`, `saveCliente`, `deleteCliente`,
`saveProveedor`, `deleteProveedor`, `saveCRM`, `deleteCRM`, `saveDoc`, `deleteDoc`,
`saveCobro`, `deleteCobro`, `saveOperario`, `deleteOperario`, `savePresupuesto`, `deletePresupuesto`

## Variables globales principales

### index.html
```js
let inventario = [], historial = [], obras = [];
let _entregaIdx = -1;  // índice del pedido en modal entrega parcial
```

### facturacion.html
```js
let facturas=[], gastos=[], gastosInternos=[], clientes=[], proveedores=[];
let cobros=[], obras=[], crmObras=[], operarios=[], presupuestos=[];
let _giIdx=-1;  // gasto interno en edición (-1=nuevo)
let _gIdx=-1;   // gasto en edición
let _fIdx=-1;   // factura en edición
```

## Funciones clave (index.html)
| Función | Descripción |
|---|---|
| `renderContrato()` | Tab Contrato de obra — carga campos inicio/fin/importe/formaPago/plazo + tabla planFac mensual |
| `guardarContrato()` | Guarda contrato + planFac[] via saveObra (incluye todos los campos de obra) |
| `ctrRenderTabla(pf)` | Renderiza tabla mensual de facturación prevista (planFac[]) |
| `confirmarEntrega()` | Entrega parcial — capturar `var idx=_entregaIdx` ANTES de cerrar modal |
| `marcarEntregado(idx, confirmados)` | Actualiza stock y marca ítems entregados |
| `cerrarModalEntregaParcial()` | Resetea `_entregaIdx = -1` |
| `cargarDatos()` | Carga inventario + historial + obras desde API |
| `renderHistorial()` | Renderiza tabla de historial filtrada |
| `toast(msg, isError)` | Muestra notificación breve |

## Funciones clave (facturacion.html)
| Función | Descripción |
|---|---|
| `renderMargen()` | Margen por obra: facturas + gastos + gastosInternos |
| `renderContabilidad()` | IVA trimestral Mod.303, resumen fiscal |
| `calcularAlertas()` | Alertas vencimientos + modelos fiscales (111/303/115/200) |
| `calcGastoInterno()` | Calcula total con IVA en modal gastos internos |
| `toggleGIObra()` | Muestra campo obra si cat='obra' o 'nominas' |
| `guardarGastoInterno()` | Guarda con base/pctIva/iva/total/vencimiento |
| `abrirModalGastoInterno(idx)` | -1=nuevo, >=0=editar |

## Categorías GastosInternos
```js
{ oficina, nominas, vehiculos, financiero, impuestos, obra, otro }
```

## IVA trimestral — fechas límite (Mod.303 + 111 + 115)
| Trimestre | Meses | Vencimiento |
|---|---|---|
| 1T | Ene–Mar | 20 abr |
| 2T | Abr–Jun | 20 jul |
| 3T | Jul–Sep | 20 oct |
| 4T | Oct–Dic | 20 ene siguiente |
| Mod.200 (IS) | — | 25 jul |

## Bugs conocidos (ya corregidos)
1. **`confirmarEntrega()`** — `_entregaIdx` se reseteaba a -1 al cerrar modal antes de usarse.
   Fix: `var idx=_entregaIdx; cerrarModal(); marcarEntregado(idx, ...)`.
2. **Carpeta Drive CRM** — `saveCRM()` con DriveApp necesitaba redespliegue (v75).
3. **`saveHtmlAsPdf` — UrlFetchApp sin permisos** (v78) — El webapp token no tenía scope `script.external_request`.
   Fix: reemplazar la función por una versión que usa Drive Advanced Service (`Drive.Files.insert` con `convert:true`) + `DriveApp.getFileById().getAs('application/pdf')`. Sin UrlFetchApp.

## Notas de despliegue Apps Script
1. Hacer cambios en editor Apps Script y guardar (Ctrl+S)
2. **Implementar → Gestionar implementaciones → lápiz editar → Versión nueva → Implementar**
3. La URL `/exec` NO cambia. Solo el código activo se actualiza.

## Estrategia para no consumir contexto (IMPORTANTE)
- Leer este `CLAUDE.md` al inicio → contexto completo sin leer los HTML grandes
- Usar **Grep** para encontrar funciones concretas: `Grep("function renderMargen", "*.html")`
- Usar **Read con offset+limit** para leer solo la sección necesaria
- Usar **Edit** para cambios puntuales (no reescribir archivos de miles de líneas)
- Evitar screenshots salvo para verificar resultado final
- Para buscar dónde está algo: `Grep("patron", path, output_mode="content", -C 5)`

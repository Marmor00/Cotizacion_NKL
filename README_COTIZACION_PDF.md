# 📄 Sistema de Cotización Formal PDF - NKL

Sistema completo para generar cotizaciones formales en formato PDF a partir de los datos de la hoja "Generador".

## 🎯 Características principales:

✅ **Webapp intuitiva** - Interfaz moderna para capturar datos del cliente
✅ **Folios únicos automáticos** - Sistema de numeración secuencial (COT-YYYYMMDD-NNNN)
✅ **Parseo inteligente** - Lee automáticamente productos de la hoja Generador
✅ **PDF profesional** - Genera documentos con formato similar al sistema CROL
✅ **Integración con carpetas** - Guarda automáticamente en la carpeta del cliente
✅ **Multi-producto** - Selecciona qué productos incluir en la cotización

---

## 📁 Estructura de archivos:

```
apps-script-project/
├── Menu_Principal.js              # Menú unificado (reemplaza onOpen individuales)
├── Cotizacion_PDF_Main.js         # Coordinador principal
├── Cotizacion_PDF_Parser.js       # Extrae productos de Generador
├── Cotizacion_PDF_Folios.js       # Sistema de folios únicos
└── Cotizacion_PDF_Webapp.html     # Interfaz web del formulario
```

---

## 🚀 Guía rápida de uso:

### 1. Configuración inicial (solo una vez):

1. Sube todos los archivos a Apps Script (ver INSTRUCCIONES_DESPLIEGUE.md)
2. **IMPORTANTE:** Elimina o comenta el `onOpen()` de `Cotizacion_PDF_Main.js` y `Generador_Carpeta.js`
3. Deja solo el `onOpen()` de `Menu_Principal.js`
4. Guarda y recarga el spreadsheet

### 2. Generar una cotización:

1. Asegúrate de que haya productos en la hoja **Generador**
2. Menú: **📄 Cotización PDF > ✨ Generar Cotización Formal**
3. Completa el formulario:
   - **Datos del Cliente:** Nombre, RFC, email, etc.
   - **Datos de Cotización:** ID de obra, vigencia, vendedor
   - **Productos:** Se cargan automáticamente, selecciona cuáles incluir
4. Clic en **✨ Generar PDF**
5. ¡Listo! El PDF se crea y guarda en la carpeta del cliente

---

## 📋 Campos del formulario:

### Datos del Cliente (obligatorios marcados con *)
- Nombre del Cliente *
- RFC
- Email
- Teléfono
- Domicilio Fiscal
- Domicilio de Entrega

### Datos de la Cotización
- **Folio:** Se genera automáticamente (ej: COT-20251211-0001)
- **Fecha:** Se toma automáticamente
- ID de Obra / Referencia *
- Contacto (default: "TODOS")
- Vigencia en días (default: 1)
- Tiempo de entrega en días (default: 0)
- Vendedor *

### Productos
- Se detectan automáticamente desde la hoja Generador
- Puedes seleccionar cuáles incluir (checkbox)
- Puedes editar el código de cada producto

---

## 🔧 Sistema de Folios:

### Formato:
```
COT-YYYYMMDD-NNNN
```

Ejemplo: `COT-20251211-0001`

### Dónde se guarda:
Los folios se almacenan en una hoja oculta llamada **Config_Folios** con:
- Contador actual
- Historial de folios generados con fecha/hora

### Ver último folio:
Menú: **📄 Cotización PDF > 🔢 Ver Último Folio**

### Resetear contador (usar con cuidado):
En Apps Script, ejecuta la función `resetearContadorFolios()`

---

## 📝 Cómo funciona el Parser:

El sistema detecta automáticamente productos en la hoja Generador:

### Tipos de productos detectados:

1. **"Datos de la Pieza"** - Puertas, ventanas, cancelería de aluminio
2. **"Medidas"** - Cristales, cubiertas, barandales

### Información extraída:
- Descripción del producto
- Categoría/Modelo
- Clave
- Precio de venta
- Importe
- (Materiales internos NO se incluyen en el PDF)

### Probar el parser:
Menú: **📄 Cotización PDF > 🧪 Probar Parser**

Luego revisa los logs: **Extensiones > Apps Script > Ver logs de ejecución**

---

## 📄 Formato del PDF generado:

El PDF incluye:

### Página 1 - Cotización:
- **Encabezado:** Datos de la empresa (NKL)
- **Título:** "Cotización a cliente [FOLIO]"
- **Info del cliente:** Datos capturados en el formulario
- **Tabla de productos:** Con columnas:
  - No.
  - Proyecto
  - Descripción
  - Código
  - U.M.
  - Cantidad
  - P.U.
  - Importe
- **Totales:** Subtotal, IVA 16%, Total

### Página 2 - Vendedor:
- Nombre del vendedor

### Ubicación del PDF:
- Se guarda en la **carpeta del cliente** (si está configurada en F2)
- Si no hay carpeta, se guarda en "Mi Unidad"

---

## ⚙️ Configuración avanzada:

### Agregar logo de la empresa:

1. Sube el logo a Google Drive
2. Obtén el ID del archivo
3. En `Cotizacion_PDF_Main.js`, busca `[LOGO]`
4. Reemplaza con:
```javascript
var logoId = "TU_ID_DEL_LOGO";
var logoBlob = DriveApp.getFileById(logoId).getBlob();
var image = cellLogo.appendImage(logoBlob);
image.setWidth(100).setHeight(50);
```

### Personalizar términos y condiciones:

Los términos están fijos en el código. Para modificarlos:

1. Edita `Cotizacion_PDF_Main.js`
2. Busca la sección donde se genera el PDF
3. Agrega después de la sección de vendedor:

```javascript
body.appendPageBreak();
body.appendParagraph("Términos y Condiciones").setHeading(DocumentApp.ParagraphHeading.HEADING2);
body.appendParagraph("1. Las cotizaciones tienen validez...");
// etc.
```

### Cambiar formato de folio:

Edita `Cotizacion_PDF_Folios.js`, función `generarFolioUnico()`:

```javascript
var folio = "PREFIJO-" + fecha + "-" + numeroFormateado;
```

---

## 🐛 Solución de problemas comunes:

### "No se encontraron productos en Generador"
- Verifica que la hoja se llame exactamente "Generador"
- Verifica que K1 tenga datos (última fila)
- Ejecuta: **🧪 Probar Parser** para ver logs

### "Error al generar folio"
- La hoja Config_Folios se crea automáticamente
- Si hay problemas, genera folio temporal basado en timestamp

### "No se puede guardar en carpeta del cliente"
- Verifica que F2 tenga hipervínculo a la carpeta
- Verifica permisos de la carpeta
- El sistema guardará en Mi Unidad como fallback

### "La webapp no se abre"
- Verifica que el archivo HTML se llame exactamente `Cotizacion_PDF_Webapp`
- Revisa permisos de Apps Script

---

## 📊 Diferencias con el sistema CROL:

### ✅ Incluido:
- Datos del cliente y cotización
- Tabla de productos con descripción
- Precios e importes
- Totales con IVA
- Folio único
- Vendedor

### ⚠️ Simplificado:
- Sucursal: Siempre "MATRIZ" (se puede personalizar)
- Días de crédito: No incluido (se puede agregar)
- No se incluyen términos y condiciones (se pueden agregar)

### ❌ No incluido (por diseño):
- Lista detallada de materiales internos
- Fórmulas de cálculo
- Medidas internas de fabricación

---

## 🔄 Próximas mejoras sugeridas:

- [ ] Agregar términos y condiciones completos
- [ ] Template más elaborado con mejor diseño
- [ ] Opción de enviar por email automáticamente
- [ ] Historial de cotizaciones en hoja separada
- [ ] Exportar en múltiples formatos (PDF, Excel)
- [ ] Integración con sistema de facturación
- [ ] Campo de descuentos por producto
- [ ] Multi-moneda (USD, EUR)

---

## 📞 Soporte:

Para reportar bugs o solicitar mejoras, revisa primero:

1. **Logs de Apps Script:** Extensiones > Apps Script > Ver logs de ejecución
2. **INSTRUCCIONES_DESPLIEGUE.md** - Guía de instalación
3. Busca mensajes con ❌ para identificar errores específicos

---

## 📜 Licencia y Créditos:

Sistema desarrollado para **Nueva Krystalum Lomelí SA de CV**

**Versión:** 1.0
**Fecha:** Diciembre 2025
**Desarrollado con:** Google Apps Script

---

## 🎉 ¡Todo listo!

Tu sistema de cotizaciones formales está configurado y funcionando.

Para empezar, ve al menú:
**📄 Cotización PDF > ✨ Generar Cotización Formal**

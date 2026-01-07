# 📄 Instrucciones para Desplegar el Sistema de Cotización PDF

## ✅ Archivos creados:

1. **Cotizacion_PDF_Main.js** - Script principal que coordina todo
2. **Cotizacion_PDF_Parser.js** - Lee y extrae productos de la hoja Generador
3. **Cotizacion_PDF_Folios.js** - Genera folios únicos automáticos
4. **Cotizacion_PDF_Webapp.html** - Interfaz web para capturar datos

---

## 📋 Pasos para el despliegue:

### 1. Abrir el editor de Apps Script

1. Abre tu archivo de Google Sheets "Cotizacion NKL"
2. Ve al menú: **Extensiones > Apps Script**
3. Se abrirá el editor de Apps Script

### 2. Crear los archivos del proyecto

#### Archivo 1: Cotizacion_PDF_Main.js
1. En el editor, haz clic en el **+** junto a "Archivos"
2. Selecciona **Script**
3. Nómbralo: `Cotizacion_PDF_Main`
4. Copia y pega todo el contenido de `apps-script-project/Cotizacion_PDF_Main.js`

#### Archivo 2: Cotizacion_PDF_Parser.js
1. Crea otro Script nuevo
2. Nómbralo: `Cotizacion_PDF_Parser`
3. Copia y pega todo el contenido de `apps-script-project/Cotizacion_PDF_Parser.js`

#### Archivo 3: Cotizacion_PDF_Folios.js
1. Crea otro Script nuevo
2. Nómbralo: `Cotizacion_PDF_Folios`
3. Copia y pega todo el contenido de `apps-script-project/Cotizacion_PDF_Folios.js`

#### Archivo 4: Cotizacion_PDF_Webapp.html
1. Haz clic en el **+** junto a "Archivos"
2. Selecciona **HTML**
3. Nómbralo: `Cotizacion_PDF_Webapp`
4. Copia y pega todo el contenido de `apps-script-project/Cotizacion_PDF_Webapp.html`

### 3. Guardar y desplegar

1. Haz clic en el icono de **💾 Guardar** (o Ctrl+S)
2. Cierra el editor de Apps Script
3. Recarga tu Google Sheets (F5 o Ctrl+R)
4. Deberías ver un nuevo menú: **📄 Cotización PDF**

---

## 🚀 Cómo usar el sistema:

### Generar una cotización:

1. Asegúrate de que la hoja **Generador** tenga productos
2. Ve al menú: **📄 Cotización PDF > ✨ Generar Cotización Formal**
3. Se abrirá una ventana con el formulario
4. Completa los datos del cliente y de la cotización
5. Revisa los productos detectados automáticamente
6. Haz clic en **✨ Generar PDF**
7. ¡Listo! El PDF se generará y guardará en la carpeta del cliente

### Ver el último folio:

1. Ve al menú: **📄 Cotización PDF > 🔢 Ver Último Folio**

### Probar el parser:

1. Ve al menú: **📄 Cotización PDF > 🧪 Probar Parser**
2. Revisa los logs: **Extensiones > Apps Script > Ver logs de ejecución**

---

## 🔧 Configuración adicional:

### Logo de la empresa (opcional):

Para agregar el logo de NKL en el PDF:

1. Sube el logo a Google Drive
2. Haz clic derecho > Obtener enlace > Asegúrate de que sea "Cualquiera con el enlace puede ver"
3. Copia el ID del archivo (está en la URL: `https://drive.google.com/file/d/ESTE_ES_EL_ID/view`)
4. En `Cotizacion_PDF_Main.js`, busca la línea que dice `[LOGO]` y reemplázala con código para insertar la imagen

### Términos y Condiciones:

Los términos y condiciones están fijos en el código. Para actualizarlos:

1. Edita `Cotizacion_PDF_Main.js`
2. Busca la sección del PDF donde se agregan los términos
3. Modifica el texto según sea necesario

---

## ⚠️ Permisos necesarios:

La primera vez que ejecutes el script, Google te pedirá permisos:

1. **Ver y administrar hojas de cálculo** - Para leer datos de Generador
2. **Ver y administrar archivos de Drive** - Para crear y guardar PDFs
3. **Ver y administrar documentos** - Para crear el documento temporal

✅ Es seguro otorgar estos permisos, solo tu cuenta tendrá acceso.

---

## 🐛 Troubleshooting:

### No aparece el menú "📄 Cotización PDF"
- Recarga la página (F5)
- Espera unos segundos (puede tardar en cargar)

### Error al parsear productos
- Verifica que la hoja se llame exactamente "Generador"
- Verifica que K1 tenga la fórmula de última fila

### El PDF no se guarda en la carpeta del cliente
- Verifica que F2 tenga el hipervínculo a la carpeta
- Verifica que tengas permisos de edición en la carpeta

### Error de permisos
- Ve a: **Extensiones > Apps Script**
- Haz clic en el icono de ⚙️ (engranaje) > Configuración del proyecto
- Verifica que los permisos estén autorizados

---

## 📞 Soporte:

Si tienes problemas:

1. Revisa los logs: **Extensiones > Apps Script > Ver logs de ejecución**
2. Busca mensajes que empiecen con ❌ para identificar errores
3. Comparte el mensaje de error para obtener ayuda

---

## 🎉 ¡Listo!

Tu sistema de cotizaciones formales está configurado y listo para usar.

**Próximas mejoras sugeridas:**
- [ ] Agregar logo de la empresa
- [ ] Términos y condiciones personalizados
- [ ] Template más elaborado con diseño profesional
- [ ] Opción de enviar por email automáticamente
- [ ] Historial de cotizaciones en una hoja separada

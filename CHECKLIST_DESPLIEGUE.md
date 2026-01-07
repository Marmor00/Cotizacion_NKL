# ✅ Checklist de Despliegue - Sistema Cotización PDF

## Pre-requisitos:
- [ ] Tienes acceso al Google Sheets "Cotizacion NKL"
- [ ] La hoja "Generador" existe y tiene datos
- [ ] Tienes permisos de edición en el spreadsheet

---

## Paso 1: Preparar archivos
- [ ] Descarga o copia todos los archivos de `apps-script-project/`
- [ ] Verifica que tengas estos 5 archivos:
  - [ ] `Menu_Principal.js`
  - [ ] `Cotizacion_PDF_Main.js`
  - [ ] `Cotizacion_PDF_Parser.js`
  - [ ] `Cotizacion_PDF_Folios.js`
  - [ ] `Cotizacion_PDF_Webapp.html`

---

## Paso 2: Abrir Apps Script
- [ ] Abre el Google Sheets
- [ ] Ve a: **Extensiones > Apps Script**
- [ ] Se abre una nueva pestaña con el editor

---

## Paso 3: Subir archivos JavaScript

### Archivo 1: Menu_Principal
- [ ] Clic en **+ (Agregar archivo) > Script**
- [ ] Nombre: `Menu_Principal`
- [ ] Pega el contenido de `Menu_Principal.js`
- [ ] Guarda (Ctrl+S)

### Archivo 2: Cotizacion_PDF_Main
- [ ] Clic en **+ > Script**
- [ ] Nombre: `Cotizacion_PDF_Main`
- [ ] Pega el contenido
- [ ] **IMPORTANTE:** Elimina o comenta la función `onOpen()` de este archivo (líneas 8-16)
- [ ] Guarda

### Archivo 3: Cotizacion_PDF_Parser
- [ ] Clic en **+ > Script**
- [ ] Nombre: `Cotizacion_PDF_Parser`
- [ ] Pega el contenido
- [ ] Guarda

### Archivo 4: Cotizacion_PDF_Folios
- [ ] Clic en **+ > Script**
- [ ] Nombre: `Cotizacion_PDF_Folios`
- [ ] Pega el contenido
- [ ] Guarda

---

## Paso 4: Subir archivo HTML

### Archivo 5: Cotizacion_PDF_Webapp
- [ ] Clic en **+ > HTML**
- [ ] Nombre: `Cotizacion_PDF_Webapp`
- [ ] Pega el contenido de `Cotizacion_PDF_Webapp.html`
- [ ] Guarda

---

## Paso 5: Eliminar onOpen duplicados

Si ya tenías el script `Generador_Carpeta.js`:

- [ ] Abre el archivo `Generador_Carpeta.js` en el editor
- [ ] Busca la función `onOpen()` (debería estar al inicio)
- [ ] Comenta o elimina esa función (líneas 6-14)
- [ ] Guarda

**Razón:** Solo puede haber UN `onOpen()` en todo el proyecto. Ahora usaremos el de `Menu_Principal.js`

---

## Paso 6: Guardar proyecto
- [ ] Haz clic en el icono **💾 Guardar proyecto**
- [ ] Opcional: Dale un nombre al proyecto (ej: "Sistema NKL Cotizaciones")

---

## Paso 7: Probar instalación

### Test 1: Recargar Sheets
- [ ] Cierra el editor de Apps Script
- [ ] Vuelve a la pestaña del Google Sheets
- [ ] Recarga la página (F5 o Ctrl+R)
- [ ] Espera 5-10 segundos

### Test 2: Verificar menús
- [ ] Deberías ver el menú: **📄 Cotización PDF**
- [ ] Deberías ver el menú: **Acciones**

### Test 3: Probar parser
- [ ] Ve a: **📄 Cotización PDF > 🧪 Probar Parser**
- [ ] Ve a: **Extensiones > Apps Script > Ver logs de ejecución**
- [ ] Deberías ver mensajes como:
  ```
  📖 Iniciando parseo de hoja Generador...
  Última fila detectada: XX
  ✅ Total de productos detectados: X
  ```

---

## Paso 8: Autorizar permisos

La primera vez que ejecutes una función:

- [ ] Google te pedirá autorizar permisos
- [ ] Haz clic en **Revisar permisos**
- [ ] Selecciona tu cuenta de Google
- [ ] Haz clic en **Avanzado**
- [ ] Haz clic en **Ir a [nombre del proyecto] (no seguro)**
- [ ] Haz clic en **Permitir**

**Permisos necesarios:**
- Ver y administrar hojas de cálculo
- Ver y administrar archivos de Drive
- Ver y administrar documentos de Google Docs

---

## Paso 9: Prueba completa

### Test final: Generar cotización
- [ ] Ve a: **📄 Cotización PDF > ✨ Generar Cotización Formal**
- [ ] Se abre la webapp
- [ ] Los productos se cargan automáticamente
- [ ] Completa los datos del cliente:
  - Nombre: "Cliente Prueba"
  - ID Obra: "PRUEBA-001"
  - Vendedor: "Tu Nombre"
- [ ] Haz clic en **✨ Generar PDF**
- [ ] Espera 10-20 segundos
- [ ] Debería mostrar: "✅ PDF generado exitosamente!"

### Verificar PDF generado
- [ ] Revisa en Google Drive (Mi Unidad o carpeta del cliente)
- [ ] Busca: "Cotización_COT-[FECHA]-0001_Cliente Prueba.pdf"
- [ ] Abre el PDF y verifica que tenga:
  - [ ] Encabezado con datos de NKL
  - [ ] Folio único
  - [ ] Datos del cliente
  - [ ] Tabla de productos
  - [ ] Totales (Subtotal, IVA, Total)
  - [ ] Nombre del vendedor

---

## Paso 10: Configuración final (opcional)

### Agregar logo de NKL
- [ ] Sube el logo a Google Drive
- [ ] Obtén el ID del archivo
- [ ] Edita `Cotizacion_PDF_Main.js`
- [ ] Busca `[LOGO]` y reemplaza según instrucciones

### Personalizar términos y condiciones
- [ ] Edita `Cotizacion_PDF_Main.js`
- [ ] Agrega términos después de la sección de vendedor
- [ ] Guarda y prueba de nuevo

---

## ✅ Checklist de verificación final:

- [ ] ✅ El menú **📄 Cotización PDF** aparece
- [ ] ✅ La webapp se abre correctamente
- [ ] ✅ Los productos se cargan desde Generador
- [ ] ✅ El folio se genera automáticamente
- [ ] ✅ El PDF se crea con formato correcto
- [ ] ✅ El PDF se guarda en Drive
- [ ] ✅ No hay errores en los logs

---

## 🎉 ¡Sistema desplegado exitosamente!

Si todos los checks están marcados, tu sistema está listo para producción.

**Próximos pasos:**
1. Capacita a los usuarios en cómo usar la webapp
2. Genera algunas cotizaciones de prueba
3. Ajusta el diseño del PDF según feedback
4. Considera agregar mejoras (logo, términos, etc.)

---

## 🐛 ¿Algo salió mal?

Si algún check falló:

1. **Revisa los logs:** Extensiones > Apps Script > Ver logs de ejecución
2. **Busca errores:** Mensajes que empiecen con ❌
3. **Verifica nombres:** Los archivos deben llamarse exactamente como se indica
4. **Confirma permisos:** Asegúrate de haber autorizado todos los permisos
5. **Consulta el README:** README_COTIZACION_PDF.md tiene más detalles

---

## 📞 Contacto para soporte:

Si necesitas ayuda, proporciona:
- Descripción del problema
- Screenshot del error
- Logs de Apps Script (copia los mensajes con ❌)

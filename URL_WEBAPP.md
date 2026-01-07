# 🔗 URL de la WebApp - Generador de Cotizaciones NKL

## ⭐ URL de acceso directo (ACTUALIZADA - v1.2):

```
https://script.google.com/macros/s/AKfycbyUx0Bq1TGqNBNLYtMJk2Jyr44ZInvrt3oG0APlqNGR9dyM8kOp7r5hkNjdAE_rx0dolQ/exec
```

**Deployment ID:** `AKfycbyUx0Bq1TGqNBNLYtMJk2Jyr44ZInvrt3oG0APlqNGR9dyM8kOp7r5hkNjdAE_rx0dolQ`

**Versión:** v1.2 - Con función de autorización de permisos

---

## 📝 URLs anteriores (descontinuadas):

~~v1.1: `https://script.google.com/macros/s/AKfycbxQxsiLcC59iS75l82r61cAyCeK5Qq5iH55NzGhNcqy8PX3zK1QlHNPY1Nf26YzGXLV6w/exec`~~
~~v1.0: `https://script.google.com/macros/s/AKfycbynwJKjBiNkzbvq5UVM_sAlBACKiV9f3TevJeKsDjqFwP6nBnyeuNGHc90ey-yXN29JIw/exec`~~

---

## 🚀 Cómo usar:

### Opción 1: URL directa (recomendada)
1. Abre el link de arriba en tu navegador
2. Autoriza permisos (solo la primera vez)
3. ¡Listo! Puedes usar la webapp

### Opción 2: Desde el menú de Google Sheets
1. Abre el archivo "Cotizacion NKL"
2. Ve al menú: **📄 Cotización PDF > ✨ Generar Cotización Formal**
3. Se abrirá como modal dentro de Sheets

---

## 📌 Compartir con el equipo:

Puedes compartir esta URL con cualquier persona que necesite generar cotizaciones.

**Importante:**
- La primera vez que accedan, deberán autorizar permisos
- Solo pueden acceder personas con permisos en el Google Sheet
- Si cambias la configuración a "Cualquier persona", no necesitarán permisos del Sheet

---

## 🔄 Actualizar deployment:

Si haces cambios en el código y quieres actualizar la webapp:

```bash
cd "c:\Users\MM\expedientes-app\NKL\Cotizacion\apps-script-project"
clasp push
clasp deploy --description "v1.1 - Descripción de cambios"
```

Luego obtén la nueva URL con:
```bash
clasp deployments
```

---

## ⚙️ Cambiar configuración de acceso:

Para cambiar quién puede acceder a la webapp:

1. Ve a **Apps Script > Desplegar > Administrar implementaciones**
2. Haz clic en el ícono de lápiz (editar) junto al deployment
3. Cambia "Quién tiene acceso"
4. Guarda

Opciones:
- **Solo yo** - Solo tú
- **Cualquier persona de [tu organización]** - Solo tu dominio de Google Workspace
- **Cualquier persona** - Cualquiera con el link

---

## 🔒 Seguridad:

- La webapp ejecuta como el usuario que la desplegó
- Los datos se leen del Google Sheet asociado
- Los PDFs se guardan en Google Drive según los permisos del usuario

---

## 📱 Usar en dispositivos móviles:

¡Sí! La webapp es responsiva y funciona en móviles y tablets.

Puedes crear un acceso directo en el inicio de tu teléfono:
- **iOS:** Safari > Compartir > Agregar a pantalla de inicio
- **Android:** Chrome > Menú > Agregar a pantalla de inicio

---

## ✅ Estado del deployment:

**Deployment ID:** @3
**Versión:** v1.2
**Fecha:** 11 de diciembre 2025
**Estado:** ✅ Activo
**Cambios:** Agregada función `autorizarPermisos()` para forzar autorización de todos los permisos necesarios

---

¿Necesitas ayuda? Revisa README_COTIZACION_PDF.md para más información.

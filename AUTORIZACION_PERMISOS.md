# 🔐 Cómo Autorizar Permisos - IMPORTANTE

## ⚠️ Problema común:

Si al usar la webapp ves este error:
```
No tienes permiso para llamar a DocumentApp.create
```

Significa que **Google no te ha pedido autorizar el permiso de Documents** todavía.

---

## ✅ Solución: Ejecutar función `autorizarPermisos()`

### Opción 1: Desde el menú de Google Sheets (MÁS FÁCIL)

1. **Abre tu Google Sheets** "Cotización NKL"
2. **Recarga la página** (F5 o Ctrl+R) para ver el menú actualizado
3. Ve al menú: **📄 Cotización PDF**
4. Haz clic en: **🔐 Autorizar Permisos (EJECUTAR PRIMERO)**
5. Google te pedirá autorizar:
   - ✅ Ver y administrar hojas de cálculo
   - ✅ Ver y administrar archivos de Drive
   - ✅ **Ver y administrar documentos de Google Docs** ← IMPORTANTE
6. Haz clic en **"Avanzado"**
7. Haz clic en **"Ir a [proyecto] (no seguro)"**
8. Revisa los permisos y haz clic en **"Permitir"**
9. Verás un mensaje de confirmación

---

### Opción 2: Desde el editor de Apps Script

1. **Abre Apps Script:**
   - Ve a: **Extensiones > Apps Script**

2. **Busca la función:**
   - En el selector de funciones (arriba), busca: `autorizarPermisos`

3. **Ejecuta la función:**
   - Haz clic en el botón **▶️ Ejecutar**

4. **Autoriza permisos:**
   - Google te mostrará una pantalla de autorización
   - Haz clic en **"Revisar permisos"**
   - Selecciona tu cuenta
   - Haz clic en **"Avanzado"**
   - Haz clic en **"Ir a [nombre del proyecto] (no seguro)"**
   - Revisa los permisos:
     - ✅ Ver y administrar hojas de cálculo
     - ✅ Ver y administrar archivos de Drive
     - ✅ **Ver y administrar documentos de Google Docs**
   - Haz clic en **"Permitir"**

5. **Verificar:**
   - Ve a: **Ver > Registros de ejecución**
   - Deberías ver:
     ```
     ✅ Permiso de Spreadsheets autorizado
     ✅ Permiso de Drive autorizado
     ✅ Permiso de Documents autorizado
     🎉 TODOS LOS PERMISOS AUTORIZADOS CORRECTAMENTE
     ```

---

## 🎯 ¿Por qué es necesario esto?

Google Apps Script necesita que autorices **explícitamente** cada permiso la primera vez que se usa.

Aunque el permiso está declarado en `appsscript.json`, Google no lo pide hasta que una función lo **use activamente**.

La función `autorizarPermisos()`:
- Crea una hoja de cálculo ✅ (fuerza permiso de Spreadsheets)
- Lee archivos de Drive ✅ (fuerza permiso de Drive)
- **Crea un documento de Google Docs** ✅ (fuerza permiso de Documents)
- Elimina el documento de prueba

---

## 📋 Después de autorizar:

1. **Recarga la webapp** (si ya la tenías abierta)
2. **Genera una cotización de prueba**
3. Ahora debería funcionar sin errores ✅

---

## ⚠️ Mensajes de seguridad de Google:

Cuando autorices, verás:
- **"Esta app no ha sido verificada"** - Es normal para scripts personales
- **"Esta app puede acceder a tus datos"** - Es normal, es TU script

**Es seguro autorizar** porque:
- Es un script que TÚ creaste
- Solo TÚ tienes acceso
- El código es de confianza (lo puedes revisar en Apps Script)

---

## 🔄 ¿Cuándo debo volver a autorizar?

Solo necesitas autorizar UNA VEZ por cuenta de Google.

Deberás volver a autorizar si:
- Cambias de cuenta de Google
- Agregas nuevos permisos en el futuro
- Revocas los permisos manualmente

---

## 📞 ¿Sigue sin funcionar?

Si después de autorizar los permisos sigue sin funcionar:

1. **Verifica que autorizaste todos los permisos:**
   - Ve a: https://myaccount.google.com/permissions
   - Busca el nombre de tu proyecto
   - Debería mostrar los 3 permisos

2. **Revoca y vuelve a autorizar:**
   - En la misma página, haz clic en tu proyecto
   - Haz clic en **"Quitar acceso"**
   - Vuelve a ejecutar `autorizarPermisos()`

3. **Revisa los logs:**
   - Ve a: **Extensiones > Apps Script > Ver registros de ejecución**
   - Busca mensajes de error

---

## ✅ Checklist de autorización:

- [ ] Ejecuté la función `autorizarPermisos()` desde el menú o Apps Script
- [ ] Vi la pantalla de autorización de Google
- [ ] Hice clic en "Avanzado" y "Permitir"
- [ ] Vi los 3 permisos autorizados en los logs
- [ ] Recarguré la webapp
- [ ] ¡Ahora funciona! 🎉

---

¿Necesitas más ayuda? Revisa README_COTIZACION_PDF.md

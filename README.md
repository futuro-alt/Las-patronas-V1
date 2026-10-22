# Las Patronas · Huevos de rancho

Web app instalable para registrar compras a granja, ventas, merma, cobranza y recordatorios de clientes. Funciona sin señal y guarda los datos en el celular donde se usa.

## Publicar en GitHub Pages

1. En github.com crea un repositorio nuevo llamado `las-patronas`, público.
2. Toca **uploading an existing file** y arrastra **todo el contenido** de esta carpeta (los archivos y las carpetas `fonts`, `icons`, `img` y `vendor`, no la carpeta `las-patronas`). Confirma con **Commit changes**.
3. Ve a **Settings → Pages**. En *Source* elige **Deploy from a branch**, rama `main` y carpeta `/ (root)`. Guarda.
4. En uno o dos minutos queda en `https://TU-USUARIO.github.io/las-patronas/`.

## Instalar en el celular

- **iPhone:** abre el link en Safari, toca **Compartir** y elige **Agregar a pantalla de inicio**.
- **Android:** abre el link en Chrome, toca el menú ⋮ y elige **Instalar app** o **Agregar a la pantalla principal**.

## Datos y respaldos

- Los datos viven solo en el celular donde se usa la app. No se suben a GitHub ni a ninguna nube.
- En **Ajustes → Hacer respaldo** se crea un archivo `.json`. Guárdalo en Archivos o en Drive, o mándalo por WhatsApp. La app recuerda hacerlo cada 7 días.
- Para pasar los datos a otro celular: instala la app ahí y usa **Ajustes → Restaurar respaldo**.
- Si se borra la app de la pantalla de inicio o se borran los datos del navegador, los datos se pierden. Haz un respaldo antes.

## Publicar una versión nueva

1. Cambia `VERSION` en `sw.js` y `VERSION_APP` en `app.js` (por ejemplo, a `1.0.1`).
2. Sube los archivos que cambiaron al repositorio.
3. Al abrir la app aparece **Hay una versión nueva → Actualizar**.

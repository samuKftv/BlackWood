# Blackwood, 1926 · Cómo publicarlo en GitHub

La web se publica en GitHub Pages y las reservas se guardan en una hoja de Google Sheets.
Las alumnas no necesitan ninguna cuenta: abren el enlace, leen el briefing y reservan.

## 1. Crear la hoja de reservas (5 minutos)

1. Crea una hoja nueva en Google Sheets. Llámala, por ejemplo, «Reparto Blackwood».
2. En el menú, entra en **Extensiones → Apps Script**.
3. Borra el código que aparece y pega todo el contenido de `Codigo.gs`.
4. En la primera línea, cambia `cambia-este-codigo` por un código que solo sepas tú.
   Es el que usarás en la web para liberar roles.
5. Guarda (icono del disquete).
6. Pulsa **Implementar → Nueva implementación**.
   - Tipo: **Aplicación web**.
   - Ejecutar como: **Yo**.
   - Quién tiene acceso: **Cualquier usuario**.
7. Pulsa **Implementar**, acepta los permisos que te pide Google y **copia la URL** que termina en `/exec`.

## 2. Conectar la web con la hoja

1. Abre `index.html` con cualquier editor de texto (o directamente en GitHub con el lápiz).
2. Busca esta línea, casi al principio del bloque `<script>`:

   ```js
   const SHEET_URL="";
   ```

3. Pega dentro de las comillas la URL que copiaste:

   ```js
   const SHEET_URL="https://script.google.com/macros/s/XXXXXXXX/exec";
   ```

## 3. Publicar en GitHub Pages

1. Sube `index.html` a tu repositorio (si ya tienes una web, puedes ponerlo en una carpeta, por ejemplo `blackwood/index.html`).
2. En el repositorio: **Settings → Pages**, y comprueba que Pages está activado en la rama principal.
3. La web quedará en `https://TU-USUARIO.github.io/TU-REPO/` (o `/blackwood/` si usaste una carpeta).

## Cómo funciona en clase

- Las alumnas eligen familia, pulsan **Reservar** en un rol libre y escriben su nombre.
- No se puede reservar un rol ocupado. La hoja comprueba las reservas una a una, así que si dos alumnas pulsan a la vez, solo entra la primera.
- El reparto se actualiza solo cada 15 segundos.
- **Solo tú puedes liberar roles.** En la pestaña Reparto escribe tu código de docente y pulsa **Activar**; aparecerán los botones «Liberar».
- Puedes ver y editar todas las reservas directamente en la hoja de Google Sheets.

## Historias de los personajes

En la pestaña **Tu personaje** cada alumna rellena la historia de su personaje y pulsa **Enviar al profesor**.
Cada envío se añade como una fila nueva en la pestaña **Historias** de la hoja de Google (se crea sola con el primer envío).
Si una alumna envía varias veces, la última fila suya es la versión más reciente.
Mientras escribe, su borrador se guarda en su navegador; también puede descargarlo en Word.

## Si algo falla

- **«No hay conexión con el reparto»**: revisa que la URL termina en `/exec` y que la implementación tiene acceso para «Cualquier usuario».
- **Has cambiado el script**: tienes que hacer **Implementar → Gestionar implementaciones → Editar → Nueva versión** para que se aplique.
- **Sin pegar la URL**, la web funciona igual pero guarda las reservas solo en el navegador de cada persona.

# Cómo convertir esto en una APK instalable

## 1. Qué instalar en tu ordenador (una sola vez)

1. **Node.js** (versión 18 o superior) → https://nodejs.org — descarga la versión LTS e instálala (siguiente, siguiente, siguiente).
2. **Android Studio** → https://developer.android.com/studio — instálalo con las opciones por defecto. Ya incluye el Android SDK y el JDK, no hace falta instalar nada más por separado.
3. La primera vez que abras Android Studio, deja que termine el asistente inicial ("SDK Manager" / "Setup Wizard") — descargará las herramientas de compilación de Android automáticamente.

## 2. Preparar el proyecto

1. Descomprime la carpeta `gestor-pozos-app` que te he dado en tu ordenador.
2. Abre una terminal (en Windows: "símbolo del sistema" o PowerShell) dentro de esa carpeta.
3. Instala las dependencias:
   ```
   npm install
   ```
4. Genera la versión web compilada:
   ```
   npm run build
   ```
5. Añade la plataforma Android:
   ```
   npx cap add android
   npx cap sync
   ```
   Esto crea una carpeta `android/` con un proyecto de Android Studio completo.

## 3. Generar el APK

1. Abre Android Studio.
2. "Open" → selecciona la carpeta `android` que se creó dentro de `gestor-pozos-app`.
3. Espera a que Android Studio indexe el proyecto (la primera vez tarda unos minutos, va descargando dependencias de Gradle).
4. Ve al menú **Build → Build Bundle(s) / APK(s) → Build APK(s)**.
5. Cuando termine, abajo a la derecha aparecerá un aviso "APK(s) generated successfully" con un enlace **locate** — ahí está tu `.apk`, normalmente en:
   ```
   android/app/build/outputs/apk/debug/app-debug.apk
   ```
6. Copia ese archivo a tu móvil (por USB, Google Drive, WhatsApp, etc.) e instálalo. Android te pedirá permitir "instalar apps de origen desconocido" la primera vez — es normal, es tu propia app, no viene de la Play Store.

## Notas

- Este APK es una app "de depuración" (debug), perfecta para uso interno del equipo. Si algún día quieres publicarla en Google Play, hay que generar una versión firmada ("release") — dímelo cuando llegue ese momento y te explico ese paso adicional.
- Cada vez que yo te dé una versión nueva de la web (cambios, mejoras), solo tienes que repetir desde el paso 2.4 (`npm run build`, `npx cap sync`) y volver a generar el APK en Android Studio — no hace falta reinstalar nada.
- Si `npx cap add android` da algún error de "Android SDK not found", abre Android Studio → More Actions → SDK Manager, y anota la ruta del SDK que aparece arriba; luego crea un archivo `local.properties` dentro de `android/` con la línea `sdk.dir=RUTA_QUE_HAS_COPIADO`.

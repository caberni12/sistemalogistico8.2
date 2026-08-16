# Seguridad de módulos

## Comportamiento aplicado

- Sin token o con sesión inválida: redirección a `index.html`.
- Con sesión válida pero sin permiso del módulo: redirección a `main.html`.
- Un módulo con `ACTIVO != SI` queda bloqueado para usuarios normales.
- Un archivo interno no registrado en la hoja `MODULOS` queda bloqueado para usuarios normales (fail-closed).
- Los administradores (`ROL = ADMIN`) pueden abrir páginas internas autenticadas.
- `main.html` vuelve a validar el token al iniciar; ya no confía únicamente en `sessionStorage`.
- Un módulo guardado como `moduloActivo` solo se restaura si continúa dentro de los módulos permitidos.

## Páginas públicas intencionales

Actualmente `auth.js` deja públicas estas páginas:

- `index.html`
- `index - copia.html`
- `webindex.html`
- `seguimiento_cliente.html`

Si alguna de ellas debe exigir login, elimínala de `AUTH_PUBLIC_PAGES` e incluye `auth.js` en su cabecera.

## Configuración necesaria en la hoja MODULOS

Cada módulo debe tener:

- `ARCHIVO`: nombre exacto del HTML, por ejemplo `orden_pedidos.html`.
- `PERMISO`: permiso no vacío, por ejemplo `pedidos`.
- `ACTIVO`: `SI` para habilitarlo.

Los permisos del usuario en la hoja `USUARIOS` deben coincidir exactamente con el campo `PERMISO` del módulo.

## Apps Script de autenticación

El archivo `auth_backend_seguro.gs` contiene la versión endurecida del backend de autenticación.

Cambios principales:

1. `verify` vuelve a leer al usuario desde `USUARIOS` antes de renovar el token.
2. Si el usuario fue desactivado, la sesión deja de ser válida.
3. Si cambian rol o permisos, el token renovado recibe los valores actuales.
4. `listarModulos` exige un token válido.
5. Las operaciones POST también refrescan la sesión contra `USUARIOS` antes de aplicar permisos.

Para que estos cambios de servidor tengan efecto, copia/actualiza este código en el proyecto Apps Script que atiende la URL configurada en `auth.js` y vuelve a desplegar la Web App. Si actualizas un despliegue existente, conserva la misma URL para no cambiar el frontend.

## Nota sobre hosting estático

Este proyecto sirve archivos HTML estáticos. El guard implementado impide el acceso normal dentro de la aplicación y redirige al usuario en el navegador, pero un servidor estático por sí solo no puede impedir que alguien descargue el archivo HTML con una petición HTTP directa si conoce la URL. Los datos y acciones sensibles deben estar protegidos también en sus APIs/Apps Script mediante token y permisos. Para impedir incluso la descarga del HTML se necesita autenticación a nivel del servidor/hosting (middleware, edge function, reverse proxy, etc.).

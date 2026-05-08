# Sistema de Gestión Servicios AS - Orden de Pedidos

## Corrección aplicada

- La hoja correcta para Orden de Pedidos es **PEDIDOS**.
- El módulo **Orden de Pedidos** lee y muestra los datos existentes de la hoja **PEDIDOS**.
- No se usa hoja `ORDEN_PEDIDOS` separada.
- Al abrir el sistema, la tabla no carga pedidos automáticamente.
- Para listar lo que está en la hoja **PEDIDOS**, se debe presionar **Sincronizar pedidos**.
- El checkbox **Mostrar todo** filtra lo ya cargado y permite que la próxima sincronización traiga todos los estados.
- Se mantiene el envío rápido, PDF A4/80mm, cantidad solicitada, ubicación igual como se muestra, código como texto y ceros adelante.

## URL configurada

https://script.google.com/macros/s/AKfycbw47SxU42yTG1yb8Tlc7H3fnhH77SZVEYSlqyTMI8xzsXEHDK0i7PDDFaj3-Y29vQo7Ng/exec

## Reemplazo necesario

1. Subir `apps_script.gs` al proyecto Apps Script.
2. Guardar y desplegar como Web App.
3. Usar `index.html` como entrada principal o abrir `orden_pedidos.html` directamente.


## Botón Cargar Lista

El módulo Orden de Pedidos incluye el botón **Cargar Lista**, que consulta directamente la hoja **PEDIDOS**, llena la tabla con la información existente y guarda el último listado en memoria local para que siga visible al recargar la página. El botón **Sincronizar pedidos** queda reservado para enviar pedidos locales pendientes y luego refrescar la vista.

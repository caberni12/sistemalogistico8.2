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

https://script.google.com/macros/s/AKfycbxDPLaKDy5LqC9US-DQcCicPDIPb0XlxnPPA-y6N1AdDvbHZPxLzM0awD-NoFTcVk8Fkw/exec

## Reemplazo necesario

1. Subir `apps_script.gs` al proyecto Apps Script.
2. Guardar y desplegar como Web App.
3. Usar `index.html` como entrada principal o abrir `orden_pedidos.html` directamente.


## Botón Cargar Lista

El módulo Orden de Pedidos incluye el botón **Cargar Lista**, que consulta directamente la hoja **PEDIDOS**, llena la tabla con la información existente y guarda el último listado en memoria local para que siga visible al recargar la página. El botón **Sincronizar pedidos** queda reservado para enviar pedidos locales pendientes y luego refrescar la vista.

## Pantalla Cliente TV

Se agregó `pantalla_cliente_tv.html`.

Funciones incluidas:

- Vista tipo televisor con pedidos en tarjetas grandes.
- Filtro por cliente, pedido, vendedor o pikeador.
- Filtro por estado.
- Sincronización automática cada 8 segundos.
- Botón de pantalla completa.
- Botón para activar o apagar la voz.
- Cuando un pedido cambia a **TERMINADO**, muestra una tarjeta emergente por unos segundos.
- El aviso por voz dice: pedido terminado, número de pedido, cliente y vendedor asociado.

Uso directo:

- `pantalla_cliente_tv.html`
- `pantalla_cliente_tv.html?cliente=NOMBRE_CLIENTE`
- `pantalla_cliente_tv.html?pedido=P-1001`
- `pantalla_cliente_tv.html?intervalo=10000`

La primera carga no anuncia pedidos que ya estaban TERMINADOS; anuncia solamente cuando detecta el cambio desde otro estado a TERMINADO.

También se ajustó el Apps Script para que el mensaje de voz de estado **TERMINADO** incluya el vendedor asociado cuando exista ese dato en la hoja PEDIDOS.


## Pantalla Cliente TV - Tema claro y loaders

El archivo `pantalla_cliente_tv.html` quedó con tema claro por defecto y loaders en los botones principales. La sincronización manual bloquea el botón mientras carga para evitar dobles consultas.

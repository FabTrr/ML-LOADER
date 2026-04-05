//CONSTANTES CON API KEYS 

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput("No data");
    }

    Logger.log("RAW: " + e.postData.contents);
    const body = JSON.parse(e.postData.contents);

    //solo se procesa la venta
    if (body.topic !== "orders" && body.topic !== "orders_v2") {
      return ContentService.createTextOutput("Topic ignorado");
    }

    let orderId = body.resource || "";
    orderId = orderId.replace("/orders/", "").replace(/\//g, "");

    Logger.log("🧾 Order ID limpio: " + orderId);

    processOrder(orderId);

    return ContentService.createTextOutput("OK");

  } catch (error) {
    Logger.log("Error en doPost: " + error.message);
    return ContentService.createTextOutput("ERROR");
  }
}

function doGet(e) {
  if (!e || !e.parameter || !e.parameter.code) {
    return ContentService.createTextOutput("OK");
  }

  const code = e.parameter.code;
  const props = PropertiesService.getScriptProperties();
  const response = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/oauth/token",
    {
      method: "post",
      payload: {
        grant_type: "authorization_code",
        client_id: ML_CLIENT_ID,
        client_secret: ML_CLIENT_SECRET,
        code: code,
        redirect_uri: ML_REDIRECT_URI
      }
    }
  );

  const data = JSON.parse(response.getContentText());

  props.setProperty("ML_ACCESS_TOKEN", data.access_token);
  props.setProperty("ML_REFRESH_TOKEN", data.refresh_token);
  props.setProperty("ML_TOKEN_EXP", (Date.now() + data.expires_in * 1000).toString());

  return ContentService.createTextOutput("Tokens guardados correctamente");
}

function getAccessToken() {
  const props = PropertiesService.getScriptProperties();

  const accessToken = props.getProperty("ML_ACCESS_TOKEN");
  const refreshToken = props.getProperty("ML_REFRESH_TOKEN");
  const exp = Number(props.getProperty("ML_TOKEN_EXP"));

  // Si todavía es válido (dejamos 1 min de margen)
  if (accessToken && exp && Date.now() < exp - 60000) {
    return accessToken;
  }
  
  if (!refreshToken) {
    throw new Error("Falta refresh_token. Reautorizar.");
  }

  // Renovar token
  const response = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/oauth/token",
    {
      method: "post",
      payload: {
        grant_type: "refresh_token",
        client_id: ML_CLIENT_ID,
        client_secret: ML_CLIENT_SECRET,
        refresh_token: refreshToken
      },
      muteHttpExceptions: true
    }
  );

  const responseText = response.getContentText();
  const data = JSON.parse(responseText);

  if (response.getResponseCode() !== 200 || !data.access_token) {
    Logger.log("Error renovando token: " + responseText);
    throw new Error("No se pudo renovar el token.");
  }

  props.setProperty("ML_ACCESS_TOKEN", data.access_token);
  props.setProperty("ML_REFRESH_TOKEN", data.refresh_token);
  props.setProperty("ML_TOKEN_EXP", (Date.now() + data.expires_in * 1000).toString());

  Logger.log("Access token renovado con éxito");
  
  return data.access_token;
}

// ==========================================
//  INSERCION EN PLANILLA GOOGLE
// ==========================================
function insertarEnPlanilla(orderData) {
  try {
    Logger.log("Iniciando inserción detallada para venta: " + orderData.id);

    if (!orderData || !orderData.buyer) return;

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const firstName = orderData.buyer?.first_name || "";
    const lastName = orderData.buyer?.last_name || "";
    const nickname = orderData.buyer?.nickname || "";
    const nombreReal = (firstName || lastName) ? `${firstName} ${lastName}`.trim() : nickname;

    let items = [];
    if (orderData.order_items && orderData.order_items.length > 0) {
      items = orderData.order_items;
    } else if (orderData.orders && orderData.orders.length > 0) {
      orderData.orders.forEach(o => {
        if (o.order_items) items = items.concat(o.order_items);
      });
    }
    
    const productosDetalle = items.length > 0 ? items.map(item => {
      const qty = item.quantity || 1;
      const title = item.item?.title || "Producto";
      const price = item.unit_price || 0;
      const subtotal = qty * price;
      const sku = item.item?.seller_sku || item.item?.id || "S/N";
      return `${qty} x ${title} SKU: ${sku} = $${subtotal}`;
    }).join(" | ") : " ERROR: No se recuperó detalle"; 

    
    let envio = "Envío a acordar"; 
    const shippingId = orderData.shipping?.id || "";
    
    if (shippingId && shippingId !== "No ID") {
      try {
        const token = getAccessToken();
        const shipRes = UrlFetchApp.fetch(`https://api.mercadolibre.com/shipments/${shippingId}`, {
          headers: { Authorization: `Bearer ${token}` },
          muteHttpExceptions: true
        });
        if (shipRes.getResponseCode() === 200) {
          const shipData = JSON.parse(shipRes.getContentText());
          const logistic = (shipData.logistic_type || "").toLowerCase();
          if (logistic.includes("flex") || logistic.includes("self_service")) {
            envio = "Mercadoenvíos FLEX *";
          } else {
            envio = "Mercadoenvíos (Colecta)";
          }
        }
      } catch (e) { envio = "MercadoEnvíos"; }
    }

    const row = sheet.getLastRow() + 1;
    
    sheet.appendRow([
      new Date(),                           // A (1) Marca temporal
      "ML_LOADER",                          // B (2) Responsable
      nombreReal,                           // C (3) Cliente
      "MercadoLibre",                       // D (4) Canal
      nickname,                             // E (5) Usuario 
      "", "", "", "", "",                  // F, G, H, I, J
      envio,                                // K (11) Envío
      productosDetalle,                     // L (12) DETALLE
      "Público",                            // M (13) Precio
      "MercadoPago",                        // N (14) Forma de pago
      "'" + "Operación # " + orderData.id,   // O (15) Observaciones
      "",                                   // P (16) Tecnico Asignado
      "", ""                                // Q, R
    ]);

  /*  const textoNota = `Pedido Nº: ${row}`;
    crearNotaEnMercadoLibre(orderData.id, textoNota);
*/
    Logger.log(`Insertado con éxito en planilla en fila ${row} y enviada nota a ML.`);

  } catch (err) {
    Logger.log("Error en insertarEnPlanilla: " + err.message);
  }
}

// FUNCIÓN DE APOYO PARA CREAR LA NOTA EN MERCADOLIBRE
/*
function crearNotaEnMercadoLibre(orderId, textoNota) {
  const token = getAccessToken();
  const url = `https://api.mercadolibre.com/orders/${orderId}/notes`;
  
  const payload = {
    "note": textoNota
  };

  const options = {
    "method": "post",
    "headers": {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    
    if (code === 200 || code === 201) {
      Logger.log(`Nota creada en MercadoLibre para orden ${orderId}`);
    } else {
      Logger.log(`Error al crear nota en ML para orden ${orderId}: Código ${code} - ${res.getContentText()}`);
    }
  } catch (e) {
    Logger.log(`Error de conexión al crear nota en ML: ${e.message}`);
  }
}
*/
function processOrder(id, intento = 1) {
  const MAX_INTENTOS = 5;
  const token = getAccessToken();

  const r = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/orders/${id}`,
    {
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true
    }
  );

  const code = r.getResponseCode();
  const bodyText = r.getContentText();

  Logger.log(`processOrder ${id} → HTTP ${code}`);

  if (code === 200) {
    return procesarOrdenReal(JSON.parse(bodyText));
  }

  if (code === 400) {
    const body = JSON.parse(bodyText);
    if (body.error === "order_belong_pack") {
      const packId = (body.cause && body.cause.length > 0) ? body.cause[0] : null;
      if (packId) {
        Logger.log(`Detectado Pack: ${packId}. Procesando...`);
        return processPack(packId, id); 
      }
    }
  }

  if (code === 404 && intento < MAX_INTENTOS) {
    Logger.log(`404 en ${id}, reintentando ${intento}...`);
    Utilities.sleep(6000);
    return processOrder(id, intento + 1);
  }
}

function processPack(packId, originalOrderId) {
  const token = getAccessToken();
  let packData = null;

  const endpoints = [
    `https://api.mercadolibre.com/packs/${packId}`,
    `https://api.mercadolibre.com/orders/packs/${packId}`
  ];

  for (let url of endpoints) {
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) {
      packData = JSON.parse(res.getContentText());
      break;
    }
  }

  let orderBase = null;
  if (packData && packData.orders && packData.orders.length > 0) {
    orderBase = packData.orders[0];
  }

  if (!packData || !orderBase || !orderBase.buyer) {
    Logger.log("Pack incompleto. Buscando detalles en la orden individual...");
    const resOrder = UrlFetchApp.fetch(`https://api.mercadolibre.com/orders/${originalOrderId}`, {
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true
    });
    if (resOrder.getResponseCode() === 200) {
      orderBase = JSON.parse(resOrder.getContentText());
    }
  }

  if (!orderBase || !orderBase.buyer) {
    Logger.log("Imposible recuperar datos del comprador para el pack " + packId);
    return;
  }

  let todosLosItems = [];
  if (packData && packData.orders) {
    packData.orders.forEach(o => {
      if (o.order_items) todosLosItems = todosLosItems.concat(o.order_items);
    });
  }
  
  if (todosLosItems.length === 0) {
    todosLosItems = orderBase.order_items || [];
  }

  const shippingConsolidado = (packData?.shipments && packData.shipments.length > 0) 
                              ? packData.shipments[0] 
                              : orderBase.shipping;

  const ordenConsolidada = {
    id: packId,
    esCarrito: true,
    buyer: orderBase.buyer,
    shipping: shippingConsolidado,
    order_items: todosLosItems,
    total_amount: packData?.orders ? packData.orders.reduce((acc, o) => acc + (o.total_amount || 0), 0) : orderBase.total_amount,
    payments: packData?.payments || orderBase.payments || []
  };

  procesarOrdenReal(ordenConsolidada);
}


function procesarOrdenReal(order) {
  if (!order || !order.buyer || !order.order_items) {
    Logger.log("Orden descartada por falta de datos esenciales.");
    return;
  }
  
  insertarEnPlanilla(order);
}

function testInsertarOrden() {
  processOrder("xxxxxxxxxxxxxxxxxx");
}

function consultarOrdenesRecientes() {
  const token = getAccessToken();
  const SELLER_ID = "31554658"; 
  const url = `https://api.mercadolibre.com/orders/search?seller=${SELLER_ID}&sort=date_desc&limit=10`;
  
  const options = {
    "method": "get",
    "headers": { "Authorization": `Bearer ${token}` },
    "muteHttpExceptions": true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      Logger.log("Error consultando órdenes: " + response.getContentText());
      return;
    }
    
    const data = JSON.parse(response.getContentText());
    const ordenes = data.results || [];
    
    if (ordenes.length === 0) return;

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    
    let idsExistentes = [];
    if (lastRow > 1) {
      idsExistentes = sheet.getRange(2, 15, lastRow - 1, 1)
                           .getValues()
                           .map(fila => {
                             let textoCelda = fila[0].toString();
                             return textoCelda.replace("Operación #", "").trim();
                           });
    }
    
    ordenes.forEach(orden => {
      const orderIdStr = orden.id.toString();
      
      if (idsExistentes.indexOf(orderIdStr) === -1) {
        Logger.log(`Nueva orden detectada: ${orderIdStr}.`);
        processOrder(orderIdStr); 
      } else {
        Logger.log(`Skip: La orden ${orderIdStr} ya existe.`);
      }
    });
    
  } catch (error) {
    Logger.log("Error en consultarOrdenesRecientes: " + error.message);
  }
}

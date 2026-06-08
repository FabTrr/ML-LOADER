// ==========================================
// 1. CONFIGURACIÓN Y CONSTANTES (CRITICAL)
// ==========================================
// =====================================
// 2. GESTIÓN DE TOKENS (AUTORIZACIÓN)
// =====================================

// Recibe el código de autorización la primera vez
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.code) return ContentService.createTextOutput("OK");
  const code = e.parameter.code;
  const props = PropertiesService.getScriptProperties();
  
  const response = UrlFetchApp.fetch("https://api.mercadolibre.com/oauth/token", {
    method: "post",
    payload: {
      grant_type: "authorization_code",
      client_id: ML_CLIENT_ID,
      client_secret: ML_CLIENT_SECRET,
      code: code,
      redirect_uri: ML_REDIRECT_URI
    },
    muteHttpExceptions: true // Agregado por seguridad
  });
  
  const resCode = response.getResponseCode();
  const resText = response.getContentText();
  
  if (resCode !== 200) {
    return ContentService.createTextOutput("Error al autorizar inicial: " + resText);
  }
  
  const data = JSON.parse(resText);
  props.setProperty("ML_ACCESS_TOKEN", data.access_token);
  props.setProperty("ML_REFRESH_TOKEN", data.refresh_token);
  props.setProperty("ML_TOKEN_EXP", (Date.now() + data.expires_in * 1000).toString());
  
  return ContentService.createTextOutput("Tokens guardados correctamente de forma inicial.");
}

// Obtiene el token activo o lo renueva si expiró (Versión Segura)
function getAccessToken() {
  const props = PropertiesService.getScriptProperties();
  const accessToken = props.getProperty("ML_ACCESS_TOKEN");
  const refreshToken = props.getProperty("ML_REFRESH_TOKEN");
  const exp = Number(props.getProperty("ML_TOKEN_EXP"));

  // 1. Si el token actual sigue siendo valido (con un margen de 1 minuto), lo usamos
  if (accessToken && exp && Date.now() < exp - 60000) {
    return accessToken;
  }

  // 2. Si no hay refresh_token guardado, lanzamos error claro
  if (!refreshToken) {
    throw new Error("Falta ML_REFRESH_TOKEN en las propiedades del script. Es necesario reautorizar la App.");
  }

  Logger.log("🔄 El Access Token expiró o está por expirar. Intentando renovar con Refresh Token...");

  // 3. Intentamos la renovacion en Mercado Libre
  const response = UrlFetchApp.fetch("https://api.mercadolibre.com/oauth/token", {
    method: "post",
    payload: {
      grant_type: "refresh_token",
      client_id: ML_CLIENT_ID,
      client_secret: ML_CLIENT_SECRET,
      refresh_token: refreshToken
    },
    muteHttpExceptions: true // Evita que Apps Script colapse si la API da error
  });

  const resCode = response.getResponseCode();
  const resText = response.getContentText();

  // 4. VALIDACIÓN CRÍTICA: Si ML no responde con 200 OK, NO pisamos las variables antiguas
  if (resCode !== 200) {
    Logger.log(`❌ Error de ML al renovar token (Código ${resCode}): ${resText}`);
    // Si el error dice explícitamente que el refresh token es inválido, lanzamos alerta
    if (resText.includes("invalid_grant")) {
      throw new Error("El Refresh Token fue revocado o expiró. Debes volver a vincular la cuenta mediante la URL de autorización.");
    }
    // Si fue un error temporal de ML (500, timeout), devolvemos el token viejo por si acaso sirve
    if (accessToken) return accessToken;
    throw new Error("Error temporal de comunicación con Mercado Libre. Reintentar más tarde.");
  }

  // 5. Si todo salió bien (200 OK), guardamos los nuevos valores seguros de que existen
  const data = JSON.parse(resText);
  
  if (data.access_token && data.refresh_token) {
    props.setProperty("ML_ACCESS_TOKEN", data.access_token);
    props.setProperty("ML_REFRESH_TOKEN", data.refresh_token);
    props.setProperty("ML_TOKEN_EXP", (Date.now() + data.expires_in * 1000).toString());
    Logger.log("✅ Token renovado con éxito.");
    return data.access_token;
  } else {
    throw new Error("La respuesta de ML fue exitosa pero no contenía los tokens esperados.");
  }
}
// ==========================================
// 4. PROCESAMIENTO DE LÓGICA
// ==========================================

function insertarEnPlanilla(orderData) {
  try {
    const shippingId = orderData.shipping?.id || "";
    // Filtro: Si no hay envío o es "A acordar", detenemos el proceso
    if (!shippingId || shippingId === "No ID") return;

    // 1. Definimos el formato de cada ítem una sola vez
    const formatItem = (i) => {
      const qty = i.quantity || 1;
      const title = i.item?.title || "Producto";
      const unitPrice = i.unit_price || 0;
      
      // Extraemos el SKU. Si el producto no tiene SKU en ML, lo dejamos en blanco
      const sku = i.item?.seller_sku ? ` SKU: ${i.item.seller_sku}` : "";
      
      // Ajustamos el formato agregando el SKU y el signo de igual para ventas simples
      if (qty >= 2) {
        return `${qty} x ${title}${sku} (${qty} x $${unitPrice})`;
      } else {
        return `${qty} x ${title}${sku} = $${unitPrice}`;
      }
    };

    // 2. Construimos el Detalle del Pedido
    let detalleFinal = "";
    if (orderData.esCarrito && orderData.pack_orders) {
      detalleFinal = orderData.order_items.map(formatItem).join("\n");
    } else {
      detalleFinal = orderData.order_items.map(formatItem).join("\n");
    }

    // 3. Consultar tipo de envío exacto (Flex vs Colecta)
    let envioTipo = "Mercadoenvíos (Colecta)";
    try {
      const token = getAccessToken();
      const shipRes = UrlFetchApp.fetch(`https://api.mercadolibre.com/shipments/${shippingId}`, {
        headers: { Authorization: `Bearer ${token}` },
        muteHttpExceptions: true
      });
      
      if (shipRes.getResponseCode() === 200) {
        const logistic = (JSON.parse(shipRes.getContentText()).logistic_type || "").toLowerCase();
        if (logistic.includes("flex") || logistic.includes("self_service")) {
          envioTipo = "Mercadoenvíos FLEX *";
        }
      }
    } catch (e) {
      Logger.log("Error al consultar logística: " + e.message);
    }

    // 4. Enviar los datos procesados al Formulario
    enviarAFormulario(orderData, detalleFinal, envioTipo);

  } catch (err) {
    Logger.log("Error en insertarEnPlanilla: " + err.message);
  }
}

function enviarAFormulario(orderData, detalleFinal, envioDetectado) {
  const form = FormApp.openById(FORM_ID);
  const items = form.getItems();
  const response = form.createResponse();

  // 1. Extraer datos del comprador
  const nombre = ((orderData.buyer.first_name || "") + " " + (orderData.buyer.last_name || "")).trim() || orderData.buyer.nickname;
  const tel = orderData.buyer.phone ? `${orderData.buyer.phone.area_code} ${orderData.buyer.phone.number}` : "";

  // 2. Extraer el RUT (Búsqueda avanzada con Endpoint de Facturación MLU + Fallbacks)
  let rutDetectado = "";
  const tokenParaProduccion = getAccessToken();
  
  // Estrategia Principal: Buscar el billing_info.id exclusivo de facturación
  let idFiscal = buscarBillingInfoId(orderData);
  if (!idFiscal && orderData.raw_pack_data) {
    idFiscal = buscarBillingInfoId(orderData.raw_pack_data);
  }
  
  if (idFiscal) {
    rutDetectado = consultarEndpointBillingInfo(idFiscal, tokenParaProduccion);
  }

  // Fallback A: Revisar el documento principal de la cuenta (buyer.document)
  if (!rutDetectado && orderData.buyer && orderData.buyer.document) {
    const tipoDoc = (orderData.buyer.document.type || "").toUpperCase();
    if (tipoDoc.includes("RUT")) {
      rutDetectado = orderData.buyer.document.number;
    }
  }

  // Fallback B: Revisar en los datos estáticos clásicos de la orden (billing)
  if (!rutDetectado && orderData.billing && orderData.billing.billing_info) {
    const info = orderData.billing.billing_info;
    const tipoDoc = (info.doc_type || "").toUpperCase();
    if (tipoDoc.includes("RUT")) {
      rutDetectado = info.doc_number;
    }
  }

  // Fallback C: Plan de rescate por Regex en los nombres/nickname
  if (!rutDetectado) {
    const datosTexto = `${orderData.buyer?.nickname || ""} ${orderData.buyer?.first_name || ""} ${orderData.buyer?.last_name || ""}`;
    const regexRUT = /R\.?U\.?T\.?[^\d]*(\d{11,12})/i;
    const match = datosTexto.match(regexRUT);
    if (match && match[1]) {
      rutDetectado = match[1];
      Logger.log(`🔍 RUT detectado por Regex en el nombre/nickname: ${rutDetectado}`);
    }
  }

  // 3. Definir qué va en Observaciones (Vital para no duplicar carritos)
  const obsTexto = orderData.esCarrito ? orderData.id : "Operación #" + orderData.id;

  // 4. Mapeo de respuestas al Formulario
  response.withItemResponse(items[0].asListItem().createResponse("Fabiana"));
  response.withItemResponse(items[1].asTextItem().createResponse(nombre));
  response.withItemResponse(items[2].asMultipleChoiceItem().createResponse("Mercadolibre"));
  response.withItemResponse(items[3].asTextItem().createResponse(orderData.buyer.nickname));
  response.withItemResponse(items[4].asTextItem().createResponse(rutDetectado));
  response.withItemResponse(items[5].asTextItem().createResponse(tel));
  response.withItemResponse(items[6].asTextItem().createResponse(orderData.shipping?.receiver_address?.address_line || ""));
  response.withItemResponse(items[7].asTextItem().createResponse(orderData.shipping?.receiver_address?.city?.name || ""));
  response.withItemResponse(items[8].asTextItem().createResponse("")); // Agencia
  response.withItemResponse(items[9].asMultipleChoiceItem().createResponse(envioDetectado));
  response.withItemResponse(items[10].asParagraphTextItem().createResponse(detalleFinal));
  response.withItemResponse(items[11].asMultipleChoiceItem().createResponse("Público"));
  response.withItemResponse(items[12].asMultipleChoiceItem().createResponse("Mercadopago"));
  response.withItemResponse(items[13].asTextItem().createResponse(obsTexto));

  // ==========================================
  // 5. CÁLCULO DE FILA Y CREACIÓN DE NOTA (Escaneo Real)
  // ==========================================
  
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  
  // Obtenemos la última fila que Sheets "cree" que tiene datos (ej: 3580 por formatos/fórmulas)
  const filaFalsa = sheet.getLastRow();
  
  // Leemos específicamente la columna 15 (Observaciones) de arriba hacia abajo
  let ultimaFilaReal = 1;
  if (filaFalsa > 0) {
    const valoresObservaciones = sheet.getRange(1, 15, filaFalsa, 1).getValues();
    
    // Escaneamos de abajo hacia arriba buscando el primer texto real (ignorando celdas vacías)
    for (let i = valoresObservaciones.length - 1; i >= 0; i--) {
      if (valoresObservaciones[i][0].toString().trim() !== "") {
        ultimaFilaReal = i + 1; // Encontramos la fila con la última venta
        break;
      }
    }
  }
  
  // Calculamos la fila exacta en la que Forms va a insertar el nuevo pedido
  const numeroPedido = ultimaFilaReal + 1; 
  const textoNota = `Pedido ${numeroPedido}`;

  // ENVIAMOS EL FORMULARIO
  response.submit();
  Logger.log(`✅ Formulario enviado. Orden: ${orderData.id} | Fila destino real: ${numeroPedido} | RUT: ${rutDetectado || "No detectado"}`);

  // Creamos la nota en Mercado Libre usando el número calculado
  try {
    if (orderData.esCarrito && orderData.pack_orders) {
      orderData.pack_orders.forEach(ordenHija => {
        crearNotaEnMercadoLibre(ordenHija.id, textoNota);
      });
    } else {
      crearNotaEnMercadoLibre(orderData.id, textoNota);
    }
  } catch (err) {
    Logger.log("Error al intentar insertar la nota en Mercado Libre: " + err.message);
  }
}

// ==========================================
// FUNCIONES DE APOYO EXCLUSIVAS PARA FACTURACIÓN
// ==========================================

function buscarBillingInfoId(obj) {
  if (!obj) return null;
  if (obj.billing_info && obj.billing_info.id) return obj.billing_info.id;
  if (obj.billing && obj.billing.billing_info && obj.billing.billing_info.id) return obj.billing.billing_info.id;
  if (obj.billing && obj.billing.id) return obj.billing.id;
  if (obj.buyer && obj.buyer.billing_info && obj.buyer.billing_info.id) return obj.buyer.billing_info.id;
  
  if (obj.orders && obj.orders.length > 0) {
    const primeraOrden = obj.orders[0];
    if (primeraOrden.billing && primeraOrden.billing.billing_info && primeraOrden.billing.billing_info.id) return primeraOrden.billing.billing_info.id;
    if (primeraOrden.billing && primeraOrden.billing.id) return primeraOrden.billing.id;
  }
  return null;
}

function consultarEndpointBillingInfo(billingInfoId, token) {
  try {
    const url = `https://api.mercadolibre.com/orders/billing-info/MLU/${billingInfoId}`;
    const options = {
      "method": "get",
      "headers": { "Authorization": `Bearer ${token}` },
      "muteHttpExceptions": true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const resData = JSON.parse(response.getContentText());
      if (resData.doc_type && resData.doc_type.toUpperCase().includes("RUT") && resData.doc_number) {
        return resData.doc_number.toString().trim();
      }
      if (resData.identification && resData.identification.type && resData.identification.type.toUpperCase().includes("RUT")) {
        return resData.identification.number.toString().trim();
      }
    }
  } catch (e) {
    Logger.log(`⚠️ Excepción en consultarEndpointBillingInfo: ${e.message}`);
  }
  return "";
}

function crearNotaEnMercadoLibre(orderId, textoNota) {
  const token = getAccessToken();
  const url = `https://api.mercadolibre.com/orders/${orderId}/notes`;
  
  const payload = { "note": textoNota };
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

// ==========================================
// FUNCIONES DE APOYO (API ML)
// ==========================================

function processOrder(id, intento = 1) {
  const token = getAccessToken();
  const r = UrlFetchApp.fetch(`https://api.mercadolibre.com/orders/${id}`, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  if (r.getResponseCode() === 200) return insertarEnPlanilla(JSON.parse(r.getContentText()));
  
  if (r.getResponseCode() === 400) {
    const body = JSON.parse(r.getContentText());
    if (body.error === "order_belong_pack") {
      const packId = body.cause?.[0];
      if (packId) return processPack(packId, id);
    }
  }
}

function processPack(packId, originalOrderId) {
  const token = getAccessToken();
  const res = UrlFetchApp.fetch(`https://api.mercadolibre.com/packs/${packId}`, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });
  
  if (res.getResponseCode() !== 200) return;
  const packData = JSON.parse(res.getContentText());
  const orderBase = packData.orders[0];

  const todosLosIds = packData.orders.map(o => `Operación #${o.id}`).join(" ");

  const ordenConsolidada = {
    id: todosLosIds,
    esCarrito: true,
    buyer: orderBase.buyer,
    billing: orderBase.billing,
    shipping: packData?.shipments?.[0] || orderBase.shipping,
    order_items: packData.orders.flatMap(o => o.order_items || []),
    pack_orders: packData.orders,
    raw_pack_data: packData // Resguardamos la estructura cruda por si el ID de facturación vive ahí
  };

  insertarEnPlanilla(ordenConsolidada);
}

function consultarOrdenesRecientes() {
  const token = getAccessToken();
  const SELLER_ID = "31554658"; 
  const url = `https://api.mercadolibre.com/orders/search?seller=${SELLER_ID}&sort=date_desc&limit=2`;
  
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
      const valoresCeldas = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
      idsExistentes = valoresCeldas.map(fila => {
        let textoCelda = fila[0].toString();
        return textoCelda.replace(/\D/g, "").trim();
      });
    }
    
    const ORDENES_SALTADAS = ["2000012958653253","2000012955238577","2000012949571057"];
    ordenes.forEach(orden => {
      const orderIdStr = orden.id.toString().replace(/\D/g, "").trim();
      
      if (ORDENES_SALTADAS.includes(orderIdStr)) {
        Logger.log(`🚫 Orden excluida manualmente: ${orderIdStr}. Saltando...`);
        return;
      }
      
      if (idsExistentes.indexOf(orderIdStr) === -1) {
        Logger.log(`🆕 Nueva orden detectada: ${orderIdStr}.`);
        processOrder(orderIdStr); 
      } else {
        Logger.log(`Skip: La orden ${orderIdStr} ya existe en los registros.`);
      }
    });
    
  } catch (error) {
    Logger.log("Error en consultarOrdenesRecientes: " + error.message);
  }
}

function cargarPedidosManuales() {
  const pedidosParaCargar = [

"2000016831487436",
"2000016831387042",
"2000016831363978",
"2000016831356766",
"2000016831183852",
"2000016831161462"

  ];

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    
    let observacionesGuardadas = [];
    if (lastRow > 1) {
      observacionesGuardadas = sheet.getRange(2, 15, lastRow - 1, 1).getValues().map(fila => fila[0].toString());
    }

    pedidosParaCargar.forEach(orderId => {
      const orderIdStr = orderId.toString().trim();
      const yaExiste = observacionesGuardadas.some(textoCelda => textoCelda.includes(orderIdStr));
      
      if (!yaExiste) {
        Logger.log(`⚙️ Ejecutando carga manual para la orden: ${orderIdStr}`);
        processOrder(orderIdStr);
        Utilities.sleep(2000); 
      } else {
        Logger.log(`⚠️ Skip: La orden ${orderIdStr} ya existe en la planilla. Se omite para no duplicar.`);
      }
    });
    
    Logger.log("✅ Proceso de carga manual finalizado.");

  } catch (error) {
    Logger.log("Error en la carga manual: " + error.message);
  }
}

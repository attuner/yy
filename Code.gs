/**
 * YAADY'S MILLET ROTI MEALS - MASTER SERVERLESS BACKEND
 * File: Code.gs
 * Version: 8.5.0 (Strict IST, Dynamic Prep Lead Times, Announcement Controls & Concurrency Hardened)
 */

const RECEIPTS_FOLDER_NAME = "Yaadys_Order_Receipts";
const REFUND_RECEIPTS_FOLDER_NAME = "Yaadys_Wallet_Refund_Receipts";
const MENU_IMAGES_FOLDER_NAME = "Yaadys_Menu_Images";
const TIMEZONE_IST = "Asia/Kolkata";

// --------------------------------------------------------------------------
// TIME & DATE SANITIZATION HELPERS (PREVENTS 1899 EPOCH BUG)
// --------------------------------------------------------------------------
function cleanSheetDateString(val) {
  if (!val) return Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd");
  if (val instanceof Date) {
    return Utilities.formatDate(val, TIMEZONE_IST, "yyyy-MM-dd");
  }
  return String(val).split("T")[0].trim();
}

function cleanSheetTimeString(val) {
  if (!val) return "Flexible";
  if (val instanceof Date) {
    return Utilities.formatDate(val, TIMEZONE_IST, "hh:mm a");
  }
  const str = String(val).replace(/^'/, "").trim();
  if (str.includes("1899") || str.includes("GMT")) {
    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, TIMEZONE_IST, "hh:mm a");
    }
  }
  return str;
}

// --------------------------------------------------------------------------
// REST API ROUTERS: doGet & doPost
// --------------------------------------------------------------------------
function doGet(e) {
  try {
    const params = e ? e.parameter : {};
    const action = params.action;

    switch (action) {
      case "getMenu":
        return sendJsonResponse(getMenuData());
      case "getCustomerProfile":
        return sendJsonResponse(getCustomerProfile(params.phone));
      case "getCustomerOrders":
        return sendJsonResponse(getCustomerOrders(params.phone));
      case "getOrderStatus":
        return sendJsonResponse(getOrderStatus(params.orderId));
      case "getWalletLedger":
        return sendJsonResponse(getWalletLedger(params.customerId));
      case "getAdminData":
        return sendJsonResponse(getAdminData(params.pin));
      case "getConfig":
        return sendJsonResponse(getConfigData());
      case "trackVisit":
        return sendJsonResponse(trackVisitor(params.isRepeat === "true"));
      case "GET_CHAT_MESSAGES":
        return sendJsonResponse(getChatMessages(params.customerId));
      case "GET_ALL_CHAT_THREADS":
        return sendJsonResponse(getAllChatThreads(params.pin));
      case "GET_ONLINE_USERS":
        return sendJsonResponse(getOnlineUsers(params.pin));
      case "GET_DAYWISE_REPORT":
        return sendJsonResponse(getDaywiseReport(params.pin, params.date));
      case "GET_ANALYTICS_REPORT":
        return sendJsonResponse(getAnalyticsReport(params.pin));
      case "GET_HOURLY_TRAFFIC_REPORT":
        return sendJsonResponse(getHourlyTrafficReport(params.pin, params.date));
      default:
        return sendJsonResponse({
          status: "SUCCESS",
          message: "Yaady's Master API Online (Strict IST Mode)",
          timestamp: Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd'T'HH:mm:ssXXX")
        });
    }
  } catch (err) {
    return sendJsonResponse({ status: "ERROR", message: err.toString() });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (t) {
    return sendJsonResponse({ status: "ERROR", message: "Server busy processing concurrent requests. Please retry in a moment." });
  }

  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action;

    switch (action) {
      case "REGISTER":
        return sendJsonResponse(registerCustomer(payload));
      case "LOGIN":
        return sendJsonResponse(loginCustomer(payload));
      case "FORGOT_PASSWORD":
        return sendJsonResponse(forgotPassword(payload));
      case "RESET_PASSWORD":
        return sendJsonResponse(resetPassword(payload));
      case "CHANGE_PASSWORD":
        return sendJsonResponse(changePassword(payload));
      case "CHECK_CART_VALIDITY":
        return sendJsonResponse(checkCartValidity(payload));
      case "SUBMIT_ORDER":
        return sendJsonResponse(submitOrder(payload));
      case "SUBMIT_WALLET_RECHARGE":
        return sendJsonResponse(submitWalletRecharge(payload));
      case "SUBMIT_WALLET_CLAIM":
        return sendJsonResponse(submitWalletClaim(payload));
      case "HEARTBEAT":
        return sendJsonResponse(recordHeartbeat(payload));
      case "SEND_CHAT_MESSAGE":
        return sendJsonResponse(sendChatMessage(payload));
      case "ADMIN_CONFIRM_PAYMENT":
        return sendJsonResponse(adminConfirmPayment(payload));
      case "ADMIN_SET_ORDER_STATUS":
        return sendJsonResponse(adminSetOrderStatus(payload));
      case "ADMIN_APPROVE_CLAIM":
        return sendJsonResponse(adminApproveClaim(payload));
      case "ADMIN_TOGGLE_MENU_STOCK":
        return sendJsonResponse(adminToggleMenuStock(payload));
      case "ADMIN_UPDATE_MENU_ITEM":
        return sendJsonResponse(adminUpdateMenuItem(payload));
      case "ADMIN_ADD_MENU_ITEM":
        return sendJsonResponse(adminAddMenuItem(payload));
      case "ADMIN_DELETE_MENU_ITEM":
        return sendJsonResponse(adminDeleteMenuItem(payload));
      case "ADMIN_UPDATE_CONFIG":
        return sendJsonResponse(adminUpdateConfig(payload));
      case "ADMIN_UPDATE_ANNOUNCEMENT":
        return sendJsonResponse(adminUpdateAnnouncement(payload));
      default:
        return sendJsonResponse({ status: "ERROR", message: "Unknown POST action: " + action });
    }
  } catch (err) {
    return sendJsonResponse({ status: "ERROR", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function sendJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------
// DATABASE SETUP & SEEDING
// -------------------------------------------------------------
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = [
    {
      name: "Customers",
      headers: ["Customer_ID", "Full_Name", "Mobile_Number", "Password_Hash", "WhatsApp_Number", "Email", "Wallet_Balance", "Total_Orders_Count", "Registered_At", "Account_Status", "Reset_Code", "Reset_Expires"]
    },
    {
      name: "Menu",
      headers: ["Item_ID", "Category", "Item_Name", "Description", "Millet_Type", "Base_Price", "Image_URL", "Is_Available", "Prep_Time_Mins", "Nutritional_Info"]
    },
    {
      name: "Orders",
      headers: ["Order_ID", "Timestamp", "Customer_ID", "Customer_Name", "Customer_Phone", "Items_JSON", "Total_Items_Count", "Subtotal", "Discount_Applied", "Discount_Slab", "Final_Payable", "Payment_Mode", "UTR_Number", "Payment_Screenshot_Drive_URL", "Pickup_Date", "Pickup_Time", "Order_Status", "Admin_Notes"]
    },
    {
      name: "Wallet_Ledger",
      headers: ["Txn_ID", "Timestamp", "Customer_ID", "Txn_Type", "Amount", "UTR_Ref", "Status", "Balance_After_Txn", "Drive_Receipt_URL", "Payout_UPI", "Withdrawal_Reason"]
    },
    {
      name: "Config",
      headers: ["Key", "Value", "Description"]
    },
    {
      name: "Analytics",
      headers: ["Date", "New_Visits", "Repeat_Visits", "Total_Visits", "New_Registrations", "Total_Orders", "Gross_Revenue"]
    },
    {
      name: "Hourly_Traffic",
      headers: ["Date", "Hour", "New_Visits", "Repeat_Visits", "Total_Visits", "Orders_Placed"]
    },
    {
      name: "Chat_Messages",
      headers: ["Message_ID", "Timestamp", "Customer_ID", "Customer_Name", "Sender", "Message_Text", "Is_Read"]
    },
    {
      name: "Active_Sessions",
      headers: ["Customer_ID", "Customer_Name", "Phone", "Last_Heartbeat"]
    }
  ];

  sheets.forEach(function(s) {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight("bold").setBackground("#EDE8DE");
      sheet.setFrozenRows(1);
    } else {
      const currentHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
      s.headers.forEach((h) => {
        if (!currentHeaders.includes(h)) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h).setFontWeight("bold").setBackground("#EDE8DE");
        }
      });
    }
  });

  const configSheet = ss.getSheetByName("Config");
  if (configSheet.getLastRow() === 1) {
    const defaultConfigs = [
      ["KITCHEN_UPI_ID", "yaadyskitchen@upi", "UPI VPA address for customer payments"],
      ["KITCHEN_NAME", "Yaady's Millet Roti Meals", "Display Brand Name"],
      ["KITCHEN_PHONE", "+919000001690", "Primary Kitchen Helpline & WhatsApp"],
      ["NOTIFICATION_EMAIL", Session.getEffectiveUser().getEmail() || "admin@yaadys.com", "Admin Alert Email recipient"],
      ["KITCHEN_ADDRESS", "Neknampur, Manikonda, Hyderabad", "Pickup Counter Address"],
      ["MIN_PREP_TIME_MINS", "15", "Minimum baseline kitchen preparation minutes"],
      ["WALLET_SYSTEM_ENABLED", "TRUE", "Master switch for wallet prepaid perks and checkout"],
      ["WALLET_SLAB_1_MIN", "500", "Tier 1 Wallet Balance minimum"],
      ["WALLET_SLAB_1_DISC", "0.10", "Tier 1 Cart Discount (10%)"],
      ["WALLET_SLAB_2_MIN", "1000", "Tier 2 Wallet Balance minimum"],
      ["WALLET_SLAB_2_DISC", "0.15", "Tier 2 Cart Discount (15%)"],
      ["WALLET_SLAB_3_MIN", "1999", "Tier 3 Wallet Balance minimum"],
      ["WALLET_SLAB_3_DISC", "0.20", "Tier 3 Cart Discount (20%)"],
      ["KITCHEN_OPEN", "TRUE", "Kitchen order acceptance master toggle"],
      ["AUTO_REFRESH_INTERVAL_SEC", "10", "Client polling interval"],
      ["ADMIN_PIN", "1234", "Master Partner Portal Access PIN"],
      ["KITCHEN_OPEN_HOUR_IST", "10:00", "Kitchen opening time in IST (HH:mm)"],
      ["KITCHEN_CLOSE_HOUR_IST", "22:00", "Kitchen closing time in IST (HH:mm)"],
      ["ANNOUNCEMENT_BANNER", "Fresh Harvest Specials: Hot Jowar & Ragi Rotis made live upon pickup!", "Broadcast Banner"]
    ];
    defaultConfigs.forEach(function(c) { configSheet.appendRow(c); });
  }

  const menuSheet = ss.getSheetByName("Menu");
  if (menuSheet.getLastRow() === 1) {
    const defaultMenu = [
      ["ITEM-101", "Millet Rotis", "Jowar Roti Thali (2 Hand-Patted Rotis)", "Traditional hand-patted sorghum rotis served with Yennegai brinjal curry, shenga chutney pudi, and fresh curd.", "Jowar", 140, "https://images.unsplash.com/photo-1626074353765-517a681e40be?w=600&auto=format&fit=crop&q=80", true, 25, "Cal: 380kcal | Protein: 9g | Fiber: 8g"],
      ["ITEM-102", "Millet Rotis", "Ragi Roti Supreme (2 Hot Rotis)", "Nutritious finger millet rotis loaded with finely chopped onions, grated carrots, coriander, and tempered cumin.", "Ragi", 130, "https://images.unsplash.com/photo-1546833999-b9f581a1996d?w=600&auto=format&fit=crop&q=80", true, 20, "Cal: 320kcal | Protein: 7g | Fiber: 11g"],
      ["ITEM-103", "Combo Meals", "Grand Millet Executive Thali", "1 Jowar Roti, 1 Bajra Roti, organic Foxtail Millet Khichdi, Dal Tadka, Seasonal Veg Curry, and chilled Spiced Buttermilk.", "Mixed Millets", 240, "https://images.unsplash.com/photo-1585937421612-70a008356fbe?w=600&auto=format&fit=crop&q=80", true, 30, "Cal: 580kcal | Protein: 16g | Fiber: 14g"],
      ["ITEM-104", "Curries & Sides", "Desi Yennegai Brinjal Curry", "Slow-cooked stuffed baby brinjals in a rich roasted peanut, white sesame, and dry-coconut gravy.", "Traditional", 110, "https://images.unsplash.com/photo-1546833998-877b37c2e5c6?w=600&auto=format&fit=crop&q=80", true, 15, "Cal: 220kcal | Protein: 4g | Fiber: 6g"],
      ["ITEM-105", "Healthy Beverages", "Earthen Claypot Spiced Buttermilk", "Churned A2 cultured yogurt infused with crushed ginger, fresh curry leaves, coriander seeds, and Himalayan rock salt.", "Probiotic", 45, "https://images.unsplash.com/photo-1556881286-fc6915169721?w=600&auto=format&fit=crop&q=80", true, 5, "Cal: 60kcal | Protein: 3g | Fiber: 0.5g"]
    ];
    defaultMenu.forEach(function(m) { menuSheet.appendRow(m); });
  }

  getOrCreateFolder(RECEIPTS_FOLDER_NAME);
  getOrCreateFolder(REFUND_RECEIPTS_FOLDER_NAME);
  getOrCreateFolder(MENU_IMAGES_FOLDER_NAME);
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  const f = DriveApp.createFolder(folderName);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f;
}

function hashString(str) {
  const rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  let hash = "";
  for (let i = 0; i < rawBytes.length; i++) {
    let byteVal = rawBytes[i];
    if (byteVal < 0) byteVal += 256;
    let hex = byteVal.toString(16);
    if (hex.length === 1) hex = "0" + hex;
    hash += hex;
  }
  return hash;
}

// -------------------------------------------------------------
// ORDERS & WALLET ENGINE (STRICT IST & ATOMIC WRITES)
// -------------------------------------------------------------
function submitOrder(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customersSheet = ss.getSheetByName("Customers");
  const ordersSheet = ss.getSheetByName("Orders");
  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
  const config = getConfigData().data;

  if (String(config.KITCHEN_OPEN).toUpperCase() !== "TRUE") {
    return { status: "ERROR", message: "The kitchen is currently closed for pickups. Orders cannot be scheduled." };
  }

  if (payload.paymentMode === "Wallet" && String(config.WALLET_SYSTEM_ENABLED).toUpperCase() !== "TRUE") {
    return { status: "ERROR", message: "Yaady Wallet payment is currently paused by admin. Please pay using Direct UPI." };
  }

  const phone = String(payload.customerPhone || "").trim();
  const custData = customersSheet.getDataRange().getValues();
  let custRowIndex = -1;
  let currentBalance = 0;
  let customerId = payload.customerId;

  for (let i = 1; i < custData.length; i++) {
    if (String(custData[i][2]).trim() === phone) {
      custRowIndex = i + 1;
      customerId = custData[i][0];
      currentBalance = Number(custData[i][6]) || 0;
      break;
    }
  }

  if (custRowIndex === -1) {
    return { status: "ERROR", message: "Customer account not found. Please log in again." };
  }

  let screenshotUrl = "";
  if (payload.base64Screenshot && payload.base64Screenshot.length > 50) {
    try {
      const folder = getOrCreateFolder(RECEIPTS_FOLDER_NAME);
      const contentType = payload.base64Screenshot.substring(payload.base64Screenshot.indexOf(":") + 1, payload.base64Screenshot.indexOf(";"));
      const rawBase64 = payload.base64Screenshot.substring(payload.base64Screenshot.indexOf(",") + 1);
      const decodedBytes = Utilities.base64Decode(rawBase64);
      const blob = Utilities.newBlob(decodedBytes, contentType, "Receipt_" + customerId + "_" + Date.now() + ".png");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      screenshotUrl = file.getUrl();
    } catch (e) {
      screenshotUrl = "Failed upload: " + e.message;
    }
  }

  const orderId = payload.preAllocatedOrderId || ("ORD-" + Math.floor(100000 + Math.random() * 900000));
  const payable = Number(payload.finalPayable) || 0;
  const subtotal = Number(payload.subtotal) || payable;
  const discountApplied = Number(payload.discountApplied) || 0;
  const itemsCount = Number(payload.totalItemsCount) || 1;
  let orderStatus = "Pending_Verification";

  const nowIst = new Date();
  const istFormattedTimestamp = Utilities.formatDate(nowIst, TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");

  if (payload.paymentMode === "Wallet") {
    if (currentBalance < payable) {
      return { status: "ERROR", message: "Insufficient wallet balance. Please top up or use Direct UPI." };
    }
    const newBalance = currentBalance - payable;
    customersSheet.getRange(custRowIndex, 7).setValue(newBalance);

    const txnId = "TXN-" + Date.now();
    ledgerSheet.appendRow([
      txnId, istFormattedTimestamp, customerId, "Debit_Order", payable,
      "Order " + orderId, "Approved", newBalance, "", "", "Meal Purchase"
    ]);

    orderStatus = "Kitchen_Accepted";
  }

  const currentOrders = Number(customersSheet.getRange(custRowIndex, 8).getValue()) || 0;
  customersSheet.getRange(custRowIndex, 8).setValue(currentOrders + 1);

  const cleanPickupDate = String(payload.pickupDate || "").split("T")[0].trim();
  const cleanPickupTime = String(payload.pickupTime || "").trim();
  const storedPickupTimeString = cleanPickupTime ? "'" + cleanPickupTime : "'Flexible";

  ordersSheet.appendRow([
    orderId,
    istFormattedTimestamp,
    customerId,
    String(payload.customerName || "Customer").trim(),
    phone,
    JSON.stringify(payload.items || []),
    itemsCount,
    subtotal,
    discountApplied,
    String(payload.discountSlab || "0%"),
    payable,
    String(payload.paymentMode || "Direct_UPI"),
    String(payload.utrNumber || (payload.paymentMode === "Wallet" ? "WALLET_DEDUCT" : "")),
    screenshotUrl,
    cleanPickupDate,
    storedPickupTimeString,
    orderStatus,
    String(payload.adminNotes || "")
  ]);

  recordOrderAnalytics(itemsCount, payable);
  recordHourlyOrderPlacement();

  try {
    sendAdminNotificationEmail("ORDER", {
      orderId: orderId,
      customerName: payload.customerName,
      customerPhone: phone,
      finalPayable: payable,
      items: payload.items || [],
      pickupDate: cleanPickupDate,
      pickupTime: cleanPickupTime,
      paymentMode: payload.paymentMode,
      utrNumber: payload.utrNumber || "N/A"
    });
  } catch (emailErr) {
    console.warn("Email alert warning: " + emailErr.message);
  }

  return {
    status: "SUCCESS",
    message: "Order placed successfully!",
    orderId: orderId,
    orderStatus: orderStatus
  };
}

function submitWalletRecharge(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customersSheet = ss.getSheetByName("Customers");
  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
  const config = getConfigData().data;

  if (String(config.WALLET_SYSTEM_ENABLED).toUpperCase() !== "TRUE") {
    return { status: "ERROR", message: "Wallet recharge is currently paused by admin." };
  }

  const phone = String(payload.phone || "").trim();
  const custData = customersSheet.getDataRange().getValues();
  let customerId = "";
  let customerName = "";
  let currentBalance = 0;

  for (let i = 1; i < custData.length; i++) {
    if (String(custData[i][2]).trim() === phone) {
      customerId = custData[i][0];
      customerName = custData[i][1];
      currentBalance = Number(custData[i][6]) || 0;
      break;
    }
  }

  if (!customerId) return { status: "ERROR", message: "Customer profile not found." };

  const txnId = String(payload.preAllocatedTxnId || ("WLT-" + Date.now())).trim();
  const amount = Number(payload.amount) || 0;
  const utr = String(payload.utrNumber || "").trim();

  if (!/^\d{12}$/.test(utr)) {
    return { status: "ERROR", message: "UTR must be a valid 12-digit numerical reference." };
  }

  const ledgerData = ledgerSheet.getDataRange().getValues();
  for (let j = 1; j < ledgerData.length; j++) {
    if (String(ledgerData[j][5]).trim() === utr) {
      return { status: "ERROR", message: "This UTR number has already been submitted." };
    }
  }

  const istFormattedTimestamp = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");

  ledgerSheet.appendRow([
    txnId, istFormattedTimestamp, customerId, "Credit_Recharge", amount,
    utr, "Pending_Verification", currentBalance, "", "", `Wallet Pack Top-Up (${txnId})`
  ]);

  try {
    sendAdminNotificationEmail("RECHARGE", {
      txnId: txnId,
      customerName: customerName,
      customerPhone: phone,
      amount: amount,
      utrNumber: utr
    });
  } catch (emailErr) {}

  return {
    status: "SUCCESS",
    message: `Recharge of ₹${amount} submitted (Token #${txnId}). Funds will reflect once kitchen verifies payment.`,
    txnId: txnId
  };
}

function sendAdminNotificationEmail(type, details) {
  const config = getConfigData().data;
  const adminEmail = config.NOTIFICATION_EMAIL || Session.getEffectiveUser().getEmail();
  if (!adminEmail || !adminEmail.includes("@")) return;

  let subject = "";
  let bodyHtml = "";

  if (type === "ORDER") {
    const token = String(details.orderId).slice(-4);
    subject = `🔔 New Order [Token #${token}] - ₹${details.finalPayable} - Yaady's Kitchen`;
    const itemsHtml = (details.items || []).map(i => `<tr><td style="padding:8px 12px;border-bottom:1px solid #eee;">${i.name}</td><td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">${i.quantity}</td><td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">₹${Number(i.lineTotal) || (Number(i.price) * Number(i.quantity))}</td></tr>`).join("");

    bodyHtml = `
      <div style="font-family:'Segoe UI',sans-serif;max-width:600px;margin:auto;border:1px solid #e0e0e0;border-radius:12px;overflow:hidden;">
        <div style="background:#1C3D2B;color:#fff;padding:20px;text-align:center;">
          <h2 style="margin:0;color:#F59E0B;">YAADY'S MILLET ROTI MEALS</h2>
          <p style="margin:5px 0 0;font-size:14px;color:#d0e0d5;">New Pickup Order Received</p>
        </div>
        <div style="padding:24px;">
          <div style="background:#FAF7F2;padding:16px;border-radius:8px;margin-bottom:20px;">
            <p style="margin:0;font-size:20px;font-weight:bold;color:#1C3D2B;">Pickup Token: #${token} (${details.orderId})</p>
            <p style="margin:6px 0 0;font-size:14px;color:#555;">Scheduled Pickup: <strong>${details.pickupDate} at ${details.pickupTime} (IST)</strong></p>
          </div>
          <table style="width:100%;font-size:14px;margin-bottom:20px;">
            <tr><td style="color:#666;">Customer Name:</td><td><strong>${details.customerName}</strong></td></tr>
            <tr><td style="color:#666;">Contact Phone:</td><td><strong>+91 ${details.customerPhone}</strong></td></tr>
            <tr><td style="color:#666;">Payment Mode:</td><td><strong>${details.paymentMode}</strong></td></tr>
            <tr><td style="color:#666;">Bank UTR / Ref:</td><td><strong>${details.utrNumber}</strong></td></tr>
          </table>
          <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">
            <thead>
              <tr style="background:#eee;"><th style="padding:8px 12px;text-align:left;">Item</th><th style="padding:8px 12px;text-align:center;">Qty</th><th style="padding:8px 12px;text-align:right;">Price</th></tr>
            </thead>
            <tbody>${itemsHtml}</tbody>
          </table>
          <div style="background:#1C3D2B;color:#fff;padding:14px 20px;border-radius:8px;display:flex;justify-content:space-between;align-items:center;">
            <span style="font-size:15px;font-weight:bold;">Final Payable Amount:</span>
            <span style="font-size:22px;font-weight:bold;color:#F59E0B;">₹${details.finalPayable}</span>
          </div>
        </div>
      </div>
    `;
  } else if (type === "RECHARGE") {
    subject = `💰 New Wallet Top-Up Request - ₹${details.amount} - Yaady's Kitchen`;
    bodyHtml = `
      <div style="font-family:'Segoe UI',sans-serif;max-width:550px;margin:auto;border:1px solid #e0e0e0;border-radius:12px;overflow:hidden;">
        <div style="background:#1C3D2B;color:#fff;padding:18px;text-align:center;">
          <h3 style="margin:0;color:#F59E0B;">YAADY'S PREPAID WALLET</h3>
          <p style="margin:4px 0 0;font-size:13px;">Customer Top-Up Awaiting Verification</p>
        </div>
        <div style="padding:20px;">
          <p style="font-size:14px;color:#333;">A customer has submitted a manual UPI top-up for wallet verification:</p>
          <div style="background:#f9f9f9;padding:14px;border-radius:8px;font-size:14px;line-height:1.8;">
            <strong>Txn Token:</strong> ${details.txnId}<br>
            <strong>Customer:</strong> ${details.customerName} (+91 ${details.customerPhone})<br>
            <strong>Recharge Amount:</strong> <span style="font-size:18px;color:#1C3D2B;font-weight:bold;">₹${details.amount}</span><br>
            <strong>Submitted 12-Digit UTR:</strong> <span style="font-family:monospace;font-weight:bold;color:#D97706;">${details.utrNumber}</span>
          </div>
        </div>
      </div>
    `;
  }

  GmailApp.sendEmail(adminEmail, subject, "", { htmlBody: bodyHtml });
}

// -------------------------------------------------------------
// HOURLY TELEMETRY RADAR ENGINE
// -------------------------------------------------------------
function trackVisitor(isRepeat) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analyticsSheet = ss.getSheetByName("Analytics");
  const hourlySheet = ss.getSheetByName("Hourly_Traffic");

  const nowIst = new Date();
  const todayStr = Utilities.formatDate(nowIst, TIMEZONE_IST, "yyyy-MM-dd");
  const currentHour = parseInt(Utilities.formatDate(nowIst, TIMEZONE_IST, "H"), 10);

  const aData = analyticsSheet.getDataRange().getValues();
  let dailyFound = false;
  for (let i = 1; i < aData.length; i++) {
    if (String(aData[i][0]) === todayStr) {
      let newVisits = Number(aData[i][1]) || 0;
      let repeatVisits = Number(aData[i][2]) || 0;
      let totalVisits = Number(aData[i][3]) || 0;

      if (isRepeat) repeatVisits += 1;
      else newVisits += 1;
      totalVisits += 1;

      analyticsSheet.getRange(i + 1, 2).setValue(newVisits);
      analyticsSheet.getRange(i + 1, 3).setValue(repeatVisits);
      analyticsSheet.getRange(i + 1, 4).setValue(totalVisits);
      dailyFound = true;
      break;
    }
  }
  if (!dailyFound) {
    analyticsSheet.appendRow([todayStr, isRepeat ? 0 : 1, isRepeat ? 1 : 0, 1, 0, 0, 0]);
  }

  const hData = hourlySheet.getDataRange().getValues();
  let hourlyFound = false;
  for (let j = 1; j < hData.length; j++) {
    if (String(hData[j][0]) === todayStr && Number(hData[j][1]) === currentHour) {
      let hNew = Number(hData[j][2]) || 0;
      let hRep = Number(hData[j][3]) || 0;
      let hTot = Number(hData[j][4]) || 0;

      if (isRepeat) hRep += 1;
      else hNew += 1;
      hTot += 1;

      hourlySheet.getRange(j + 1, 3).setValue(hNew);
      hourlySheet.getRange(j + 1, 4).setValue(hRep);
      hourlySheet.getRange(j + 1, 5).setValue(hTot);
      hourlyFound = true;
      break;
    }
  }
  if (!hourlyFound) {
    hourlySheet.appendRow([todayStr, currentHour, isRepeat ? 0 : 1, isRepeat ? 1 : 0, 1, 0]);
  }

  return { status: "SUCCESS", date: todayStr, hour: currentHour, isRepeat: isRepeat };
}

function recordHourlyOrderPlacement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hourlySheet = ss.getSheetByName("Hourly_Traffic");
  const nowIst = new Date();
  const todayStr = Utilities.formatDate(nowIst, TIMEZONE_IST, "yyyy-MM-dd");
  const currentHour = parseInt(Utilities.formatDate(nowIst, TIMEZONE_IST, "H"), 10);

  const hData = hourlySheet.getDataRange().getValues();
  for (let j = 1; j < hData.length; j++) {
    if (String(hData[j][0]) === todayStr && Number(hData[j][1]) === currentHour) {
      const currentOrders = Number(hData[j][5]) || 0;
      hourlySheet.getRange(j + 1, 6).setValue(currentOrders + 1);
      return;
    }
  }
  hourlySheet.appendRow([todayStr, currentHour, 0, 0, 0, 1]);
}

function getHourlyTrafficReport(pin, targetDateStr) {
  const config = getConfigData().data;
  if (String(pin) !== String(config.ADMIN_PIN)) {
    return { status: "ERROR", message: "Unauthorized PIN." };
  }

  const queryDate = targetDateStr || Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hourlySheet = ss.getSheetByName("Hourly_Traffic");
  const hData = hourlySheet.getDataRange().getValues();

  const hourlyMap = {};
  for (let h = 0; h < 24; h++) {
    const startHour12 = (h % 12 === 0 ? 12 : h % 12);
    const startPeriod = h < 12 ? "AM" : "PM";
    const nextH = (h + 1) % 24;
    const endHour12 = (nextH % 12 === 0 ? 12 : nextH % 12);
    const endPeriod = (h + 1) < 12 || (h + 1) === 24 ? "AM" : "PM";

    const label = `${String(startHour12).padStart(2, '0')}:00 ${startPeriod} - ${String(endHour12).padStart(2, '0')}:00 ${endPeriod}`;

    hourlyMap[h] = {
      hour: h,
      label: label,
      newVisits: 0,
      repeatVisits: 0,
      totalVisits: 0,
      ordersPlaced: 0
    };
  }

  for (let i = 1; i < hData.length; i++) {
    if (String(hData[i][0]) === queryDate) {
      const h = Number(hData[i][1]);
      if (hourlyMap[h]) {
        hourlyMap[h].newVisits += Number(hData[i][2]) || 0;
        hourlyMap[h].repeatVisits += Number(hData[i][3]) || 0;
        hourlyMap[h].totalVisits += Number(hData[i][4]) || 0;
        hourlyMap[h].ordersPlaced += Number(hData[i][5]) || 0;
      }
    }
  }

  return {
    status: "SUCCESS",
    date: queryDate,
    timezone: TIMEZONE_IST,
    hourlyData: Object.values(hourlyMap)
  };
}

function getDaywiseReport(pin, targetDateStr) {
  const config = getConfigData().data;
  if (String(pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };

  const queryDate = targetDateStr || Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd");
  const ordersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  const data = ordersSheet.getDataRange().getValues();

  let totalOrders = 0, accepted = 0, preparing = 0, ready = 0, completed = 0, cancelled = 0, grossRevenue = 0, upiRevenue = 0, walletRevenue = 0, totalRotis = 0;
  const filteredOrders = [];

  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][14];
    const orderDate = cleanSheetDateString(rawDate);

    if (orderDate === queryDate) {
      const status = String(data[i][16]).trim();
      const amount = Number(data[i][10]) || 0;
      const mode = String(data[i][11]).trim();
      const itemsCount = Number(data[i][6]) || 0;

      totalOrders++;
      if (status === "Kitchen_Accepted") accepted++;
      else if (status === "Preparing") preparing++;
      else if (status === "Ready_For_Pickup") ready++;
      else if (status === "Completed") completed++;
      else if (status === "Rejected" || status === "Cancelled") cancelled++;

      if (status !== "Rejected" && status !== "Cancelled") {
        grossRevenue += amount;
        if (mode === "Wallet") walletRevenue += amount;
        else upiRevenue += amount;
        totalRotis += itemsCount;
      }

      filteredOrders.push({
        rowIndex: i + 1,
        orderId: String(data[i][0]).trim(),
        timestamp: data[i][1],
        customerName: String(data[i][3]).trim(),
        customerPhone: String(data[i][4]).trim(),
        items: JSON.parse(data[i][5] || "[]"),
        itemsCount: itemsCount,
        finalPayable: amount,
        paymentMode: mode,
        utr: String(data[i][12] || "").trim(),
        pickupDate: orderDate,
        pickupTime: cleanSheetTimeString(data[i][15]),
        orderStatus: status
      });
    }
  }

  return {
    status: "SUCCESS",
    date: queryDate,
    metrics: { totalOrders, accepted, preparing, ready, completed, cancelled, grossRevenue, upiRevenue, walletRevenue, rotisPrepared: totalRotis },
    orders: filteredOrders
  };
}

// -------------------------------------------------------------
// CUSTOMER SERVICES
// -------------------------------------------------------------
function getCustomerOrders(phone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();
  const orders = [];
  const cleanPhone = String(phone || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim() === cleanPhone) {
      const cleanDate = cleanSheetDateString(data[i][14]);
      const cleanTime = cleanSheetTimeString(data[i][15]);

      orders.push({
        orderId: String(data[i][0]).trim(),
        timestamp: data[i][1],
        customerId: String(data[i][2]).trim(),
        customerName: String(data[i][3]).trim(),
        customerPhone: String(data[i][4]).trim(),
        items: JSON.parse(data[i][5] || "[]"),
        totalItemsCount: Number(data[i][6]) || 0,
        subtotal: Number(data[i][7]) || 0,
        discountApplied: Number(data[i][8]) || 0,
        discountSlab: String(data[i][9] || ""),
        finalPayable: Number(data[i][10]) || 0,
        paymentMode: String(data[i][11] || ""),
        utrNumber: String(data[i][12] || ""),
        pickupDate: cleanDate,
        pickupTime: cleanTime,
        orderStatus: String(data[i][16] || "").trim()
      });
    }
  }
  return { status: "SUCCESS", orders: orders.reverse() };
}

function getOrderStatus(orderId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();
  const cleanOrderId = String(orderId || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === cleanOrderId) {
      const cleanDate = cleanSheetDateString(data[i][14]);
      const cleanTime = cleanSheetTimeString(data[i][15]);

      return {
        status: "SUCCESS",
        order: {
          orderId: String(data[i][0]).trim(),
          timestamp: data[i][1],
          customerName: String(data[i][3]).trim(),
          customerPhone: String(data[i][4]).trim(),
          items: JSON.parse(data[i][5] || "[]"),
          finalPayable: Number(data[i][10]) || 0,
          paymentMode: String(data[i][11] || ""),
          utrNumber: String(data[i][12] || ""),
          pickupDate: cleanDate,
          pickupTime: cleanTime,
          orderStatus: String(data[i][16] || "").trim()
        }
      };
    }
  }
  return { status: "ERROR", message: "Order not found." };
}

function getAnalyticsReport(pin) {
  const config = getConfigData().data;
  if (String(pin) !== String(config.ADMIN_PIN)) {
    return { status: "ERROR", message: "Unauthorized PIN." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analyticsSheet = ss.getSheetByName("Analytics");
  const aData = analyticsSheet.getDataRange().getValues();

  const customersSheet = ss.getSheetByName("Customers");
  const totalRegisteredUsers = Math.max(0, customersSheet.getLastRow() - 1);

  const dayWise = [];
  const weekWiseMap = {};

  for (let i = 1; i < aData.length; i++) {
    const dateStr = String(aData[i][0]);
    const newV = Number(aData[i][1]) || 0;
    const repV = Number(aData[i][2]) || 0;
    const totV = Number(aData[i][3]) || (newV + repV);
    const newReg = Number(aData[i][4]) || 0;
    const totOrders = Number(aData[i][5]) || 0;
    const grossRev = Number(aData[i][6]) || 0;

    dayWise.push({
      date: dateStr,
      newVisits: newV,
      repeatVisits: repV,
      totalVisits: totV,
      newRegistrations: newReg,
      orders: totOrders,
      revenue: grossRev
    });

    const d = new Date(dateStr + "T00:00:00+05:30");
    if (!isNaN(d.getTime())) {
      const startOfYear = new Date(d.getFullYear(), 0, 1);
      const weekNum = Math.ceil((((d - startOfYear) / 86400000) + startOfYear.getDay() + 1) / 7);
      const weekKey = `${d.getFullYear()} - W${weekNum}`;

      if (!weekWiseMap[weekKey]) {
        weekWiseMap[weekKey] = {
          week: weekKey,
          newVisits: 0,
          repeatVisits: 0,
          totalVisits: 0,
          newRegistrations: 0,
          orders: 0,
          revenue: 0
        };
      }
      weekWiseMap[weekKey].newVisits += newV;
      weekWiseMap[weekKey].repeatVisits += repV;
      weekWiseMap[weekKey].totalVisits += totV;
      weekWiseMap[weekKey].newRegistrations += newReg;
      weekWiseMap[weekKey].orders += totOrders;
      weekWiseMap[weekKey].revenue += grossRev;
    }
  }

  return {
    status: "SUCCESS",
    totalRegisteredUsers: totalRegisteredUsers,
    dayWise: dayWise.reverse(),
    weekWise: Object.values(weekWiseMap).reverse()
  };
}

function registerCustomer(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const phone = String(payload.phone || "").trim();
  const password = String(payload.password || "").trim();

  if (!/^[6-9]\d{9}$/.test(phone)) return { status: "ERROR", message: "Please enter a valid 10-digit Indian Mobile Number." };
  if (!password || password.length < 4) return { status: "ERROR", message: "Password must be at least 4 characters long." };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === phone) return { status: "ERROR", message: "Account already exists with this mobile number. Please log in." };
  }

  const randomHex = Math.floor(Math.random() * 0xFFFFFF).toString(16).padEnd(6, '0').toUpperCase();
  const customerId = "YMR-" + phone.slice(-4) + "-" + randomHex;
  const istTimestamp = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");

  sheet.appendRow([
    customerId, payload.fullName, phone, hashString(password),
    payload.whatsapp || phone, payload.email || "", 0, 0, istTimestamp, "Active", "", ""
  ]);

  recordRegistrationAnalytics();

  return {
    status: "SUCCESS",
    message: "Registration completed successfully.",
    customer: { customerId, fullName: payload.fullName, phone, whatsapp: payload.whatsapp || phone, email: payload.email || "", walletBalance: 0, totalOrders: 0 }
  };
}

function recordRegistrationAnalytics() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analytics");
  const todayStr = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === todayStr) {
      sheet.getRange(i + 1, 5).setValue((Number(data[i][4]) || 0) + 1);
      return;
    }
  }
  sheet.appendRow([todayStr, 0, 0, 0, 1, 0, 0]);
}

function recordOrderAnalytics(rotisCount, amount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analytics");
  const todayStr = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === todayStr) {
      sheet.getRange(i + 1, 6).setValue((Number(data[i][5]) || 0) + 1);
      sheet.getRange(i + 1, 7).setValue((Number(data[i][6]) || 0) + amount);
      return;
    }
  }
  sheet.appendRow([todayStr, 0, 0, 0, 0, 1, amount]);
}

function loginCustomer(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const phone = String(payload.phone || "").trim();
  const inputHash = hashString(String(payload.password || "").trim());

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === phone) {
      if (String(data[i][3]) === inputHash) {
        return {
          status: "SUCCESS",
          message: "Welcome back!",
          customer: { customerId: data[i][0], fullName: data[i][1], phone: data[i][2], whatsapp: data[i][4], email: data[i][5], walletBalance: Number(data[i][6]) || 0, totalOrders: Number(data[i][7]) || 0 }
        };
      } else {
        return { status: "ERROR", message: "Incorrect password. Please verify and try again." };
      }
    }
  }
  return { status: "ERROR", message: "Mobile number not registered. Please create an account." };
}

function getCustomerProfile(phone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const cleanPhone = String(phone || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === cleanPhone) {
      return {
        status: "SUCCESS",
        customer: { customerId: data[i][0], fullName: data[i][1], phone: data[i][2], whatsapp: data[i][4], email: data[i][5], walletBalance: Number(data[i][6]) || 0, totalOrders: Number(data[i][7]) || 0, accountStatus: data[i][9] }
      };
    }
  }
  return { status: "ERROR", message: "Customer profile not found." };
}

function getMenuData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const data = sheet.getDataRange().getValues();
  const menu = [];

  for (let i = 1; i < data.length; i++) {
    menu.push({
      id: data[i][0],
      category: data[i][1],
      name: data[i][2],
      description: data[i][3],
      milletType: data[i][4],
      price: Number(data[i][5]) || 0,
      image: data[i][6],
      available: (data[i][7] === true || String(data[i][7]).toUpperCase() === "TRUE"),
      prepTimeMins: Number(data[i][8]) || 15,
      nutrition: data[i][9]
    });
  }
  return { status: "SUCCESS", menu: menu };
}

function getConfigData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }
  return { status: "SUCCESS", data: config };
}

function submitWalletClaim(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customersSheet = ss.getSheetByName("Customers");
  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");

  const phone = String(payload.phone || "").trim();
  const claimAmount = Number(payload.amount) || 0;
  const payoutUpi = String(payload.payoutUpi || "").trim();
  const reason = String(payload.reason || "Customer refund request").trim();

  if (!payoutUpi.includes("@")) return { status: "ERROR", message: "Please provide a valid UPI ID (e.g., name@okaxis)." };

  const custData = customersSheet.getDataRange().getValues();
  let custRowIndex = -1;
  let customerId = "";
  let currentBalance = 0;

  for (let i = 1; i < custData.length; i++) {
    if (String(custData[i][2]).trim() === phone) {
      custRowIndex = i + 1;
      customerId = custData[i][0];
      currentBalance = Number(custData[i][6]) || 0;
      break;
    }
  }

  if (custRowIndex === -1) return { status: "ERROR", message: "Customer not found." };
  if (claimAmount <= 0 || claimAmount > currentBalance) {
    return { status: "ERROR", message: "Claim amount cannot exceed balance (₹" + currentBalance + ")." };
  }

  const remainingBal = currentBalance - claimAmount;
  customersSheet.getRange(custRowIndex, 7).setValue(remainingBal);

  const claimId = "CLM-" + Math.floor(10000 + Math.random() * 90000);
  const istTimestamp = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");

  ledgerSheet.appendRow([claimId, istTimestamp, customerId, "Debit_Claim", claimAmount, "HOLD", "CLAIM_PENDING", remainingBal, "", payoutUpi, reason]);

  return { status: "SUCCESS", message: "Claim logged. Payout to " + payoutUpi + " will be reviewed shortly.", claimId: claimId, newBalance: remainingBal };
}

function getWalletLedger(customerId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Wallet_Ledger");
  const data = sheet.getDataRange().getValues();
  const txns = [];
  const cleanId = String(customerId || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === cleanId) {
      txns.push({
        txnId: data[i][0],
        timestamp: data[i][1],
        type: data[i][3],
        amount: Number(data[i][4]) || 0,
        utrRef: data[i][5],
        status: data[i][6],
        balanceAfter: Number(data[i][7]) || 0,
        receiptPdfUrl: data[i][8],
        payoutUpi: data[i][9],
        reason: data[i][10]
      });
    }
  }
  return { status: "SUCCESS", transactions: txns.reverse() };
}

function getAdminData(pin) {
  const config = getConfigData().data;
  if (String(pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Invalid Kitchen Partner PIN." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const oData = ordersSheet.getDataRange().getValues();
  const orders = [];

  for (let i = 1; i < oData.length; i++) {
    const cleanDate = cleanSheetDateString(oData[i][14]);
    const cleanTime = cleanSheetTimeString(oData[i][15]);

    orders.push({
      rowIndex: i + 1,
      orderId: String(oData[i][0]).trim(),
      timestamp: oData[i][1],
      customerId: String(oData[i][2]).trim(),
      customerName: String(oData[i][3]).trim(),
      customerPhone: String(oData[i][4]).trim(),
      items: JSON.parse(oData[i][5] || "[]"),
      totalItems: Number(oData[i][6]) || 0,
      finalPayable: Number(oData[i][10]) || 0,
      paymentMode: String(oData[i][11] || ""),
      utr: String(oData[i][12] || "").trim(),
      screenshotUrl: String(oData[i][13] || ""),
      pickupDate: cleanDate,
      pickupTime: cleanTime,
      orderStatus: String(oData[i][16] || "").trim()
    });
  }

  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
  const lData = ledgerSheet.getDataRange().getValues();
  const pendingRecharges = [];
  const pendingClaims = [];

  for (let j = 1; j < lData.length; j++) {
    if (lData[j][6] === "Pending_Verification") {
      pendingRecharges.push({
        rowIndex: j + 1, txnId: lData[j][0], timestamp: lData[j][1],
        customerId: lData[j][2], amount: Number(lData[j][4]) || 0, utr: lData[j][5]
      });
    } else if (lData[j][6] === "CLAIM_PENDING") {
      pendingClaims.push({
        rowIndex: j + 1, claimId: lData[j][0], timestamp: lData[j][1],
        customerId: lData[j][2], amount: Number(lData[j][4]) || 0, payoutUpi: lData[j][9], reason: lData[j][10]
      });
    }
  }

  return {
    status: "SUCCESS",
    config: config,
    orders: orders.reverse(),
    pendingRecharges: pendingRecharges,
    pendingClaims: pendingClaims
  };
}

function adminConfirmPayment(payload) {
  const config = getConfigData().data;
  if (String(payload.pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (payload.type === "RECHARGE") {
    const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
    const customersSheet = ss.getSheetByName("Customers");
    const ledgerRow = Number(payload.rowIndex);
    const customerId = ledgerSheet.getRange(ledgerRow, 3).getValue();
    const amount = Number(ledgerSheet.getRange(ledgerRow, 5).getValue()) || 0;

    const custData = customersSheet.getDataRange().getValues();
    for (let i = 1; i < custData.length; i++) {
      if (custData[i][0] === customerId) {
        const newBal = (Number(custData[i][6]) || 0) + amount;
        customersSheet.getRange(i + 1, 7).setValue(newBal);
        ledgerSheet.getRange(ledgerRow, 7).setValue("Approved");
        ledgerSheet.getRange(ledgerRow, 8).setValue(newBal);
        return { status: "SUCCESS", message: "Wallet recharge credited & approved." };
      }
    }
    return { status: "ERROR", message: "Customer ID not matched." };
  } else {
    const ordersSheet = ss.getSheetByName("Orders");
    ordersSheet.getRange(Number(payload.rowIndex), 17).setValue("Kitchen_Accepted");
    return { status: "SUCCESS", message: "Payment verified. Order accepted." };
  }
}

function adminSetOrderStatus(payload) {
  const config = getConfigData().data;
  if (String(payload.pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };
  const ordersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  ordersSheet.getRange(Number(payload.rowIndex), 17).setValue(payload.newStatus);
  return { status: "SUCCESS", message: "Order updated to " + payload.newStatus };
}

function adminApproveClaim(payload) {
  const config = getConfigData().data;
  if (String(payload.pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
  const customersSheet = ss.getSheetByName("Customers");
  const rowIndex = Number(payload.rowIndex);
  const transferUtr = String(payload.transferUtr || "").trim();

  if (!/^\d{12}$/.test(transferUtr)) return { status: "ERROR", message: "Please provide a valid 12-digit bank transfer UTR." };

  const claimId = ledgerSheet.getRange(rowIndex, 1).getValue();
  const customerId = ledgerSheet.getRange(rowIndex, 3).getValue();
  const amount = ledgerSheet.getRange(rowIndex, 5).getValue();
  const payoutUpi = ledgerSheet.getRange(rowIndex, 10).getValue();

  let customerName = "Customer";
  let customerPhone = "";
  const custData = customersSheet.getDataRange().getValues();
  for (let i = 1; i < custData.length; i++) {
    if (custData[i][0] === customerId) {
      customerName = custData[i][1];
      customerPhone = custData[i][2];
      break;
    }
  }

  const pdfUrl = generateRefundPdf({
    claimId: claimId,
    customerName: customerName,
    customerPhone: customerPhone,
    amount: amount,
    payoutUpi: payoutUpi,
    transferUtr: transferUtr,
    timestamp: new Date()
  });

  ledgerSheet.getRange(rowIndex, 6).setValue(transferUtr);
  ledgerSheet.getRange(rowIndex, 7).setValue("Approved");
  ledgerSheet.getRange(rowIndex, 9).setValue(pdfUrl);

  return { status: "SUCCESS", message: "Claim disbursed. PDF acknowledgment generated.", pdfUrl: pdfUrl };
}

function generateRefundPdf(data) {
  const folder = getOrCreateFolder(REFUND_RECEIPTS_FOLDER_NAME);
  const formattedDate = Utilities.formatDate(data.timestamp, TIMEZONE_IST, "dd MMM yyyy, hh:mm a");

  const html = `
    <!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
      body{font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;color:#1c3d2b;padding:40px;margin:0;}
      .header{border-bottom:3px solid #1c3d2b;padding-bottom:15px;margin-bottom:25px;}
      .title{font-size:24px;font-weight:bold;margin:0;color:#1c3d2b;}
      .grid{width:100%;border-collapse:collapse;margin-top:20px;}
      .grid td{padding:10px 12px;font-size:13px;border-bottom:1px solid #eee;}
      .grid td.lbl{font-weight:bold;color:#555;width:40%;background:#faf7f2;}
      .amt-box{margin-top:25px;background:#1c3d2b;color:#fff;padding:18px;border-radius:8px;text-align:center;}
      .amt-val{font-size:26px;font-weight:bold;color:#f59e0b;margin-top:4px;}
    </style></head><body>
      <div class="header">
        <div class="title">YAADY'S MILLET ROTI MEALS</div>
        <div style="font-size:13px;color:#666;margin-top:4px;">Official Balance Refund Settlement Acknowledgment (Strict IST)</div>
      </div>
      <table class="grid">
        <tr><td class="lbl">Claim ID</td><td><strong>${data.claimId}</strong></td></tr>
        <tr><td class="lbl">Disbursement Time</td><td>${formattedDate} (IST)</td></tr>
        <tr><td class="lbl">Beneficiary Name</td><td>${data.customerName}</td></tr>
        <tr><td class="lbl">Customer Phone</td><td>+91 ${data.customerPhone}</td></tr>
        <tr><td class="lbl">Settlement UPI ID</td><td><strong>${data.payoutUpi}</strong></td></tr>
        <tr><td class="lbl">Bank Transfer UTR</td><td><strong>${data.transferUtr}</strong></td></tr>
        <tr><td class="lbl">Status</td><td style="color:#2e7d32;font-weight:bold;">DISBURSED & VERIFIED</td></tr>
      </table>
      <div class="amt-box">
        <div style="font-size:12px;letter-spacing:1px;text-transform:uppercase;">Settled Refund Amount</div>
        <div class="amt-val">₹${Number(data.amount).toFixed(2)}</div>
      </div>
    </body></html>
  `;

  const blob = Utilities.newBlob(html, "text/html", "Claim_" + data.claimId + ".html")
    .getAs("application/pdf")
    .setName("Refund_Receipt_" + data.claimId + ".pdf");
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function checkCartValidity(payload) {
  const config = getConfigData().data;
  if (String(config.KITCHEN_OPEN).toUpperCase() !== "TRUE") {
    return { status: "KITCHEN_CLOSED", message: "The kitchen is currently closed for pickups. Please check back later." };
  }

  const menuSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const menuData = menuSheet.getDataRange().getValues();
  const cartItems = payload.cartItems || [];

  const unavailableItems = [];
  const validItems = [];

  cartItems.forEach(cartItem => {
    let found = false;
    for (let i = 1; i < menuData.length; i++) {
      if (String(menuData[i][0]).trim() === String(cartItem.id).trim()) {
        found = true;
        const isAvailable = (menuData[i][7] === true || String(menuData[i][7]).toUpperCase() === "TRUE");
        if (isAvailable) {
          validItems.push({ id: menuData[i][0], name: menuData[i][2], price: Number(menuData[i][5]) || 0, quantity: cartItem.quantity });
        } else {
          unavailableItems.push({ id: menuData[i][0], name: menuData[i][2] });
        }
        break;
      }
    }
    if (!found) unavailableItems.push({ id: cartItem.id, name: "Discontinued Meal" });
  });

  return {
    status: "SUCCESS",
    kitchenOpen: true,
    hasSoldOutItems: unavailableItems.length > 0,
    unavailableItems: unavailableItems,
    validItems: validItems
  };
}

function forgotPassword(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const phone = String(payload.phone || "").trim();
  const email = String(payload.email || "").trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === phone) {
      const storedEmail = String(data[i][5] || "").trim().toLowerCase();
      if (!storedEmail || storedEmail !== email) {
        return { status: "ERROR", message: "Registered email address does not match this mobile number." };
      }

      const resetCode = Math.random().toString(36).substring(2, 8).toUpperCase();
      const expires = new Date(Date.now() + 15 * 60 * 1000);

      sheet.getRange(i + 1, 11).setValue(resetCode);
      sheet.getRange(i + 1, 12).setValue(expires);

      try {
        GmailApp.sendEmail(
          storedEmail,
          "Password Reset Code - Yaady's Millet Roti Meals",
          "",
          {
            htmlBody: `
              <div style="font-family:'Segoe UI',sans-serif;max-width:500px;margin:auto;border:1px solid #e0e0e0;border-radius:12px;padding:24px;">
                <h2 style="color:#1C3D2B;margin-top:0;">Yaady's Password Assistance</h2>
                <p>Hello <strong>${data[i][1]}</strong>,</p>
                <p>Here is your 6-character password verification code:</p>
                <div style="background:#FAF7F2;padding:16px;border-radius:8px;text-align:center;font-size:28px;font-weight:bold;letter-spacing:4px;color:#1C3D2B;border:1px solid #EDE8DE;">
                  ${resetCode}
                </div>
                <p style="font-size:12px;color:#777;margin-top:16px;">Expires in 15 minutes. If you did not request this, please ignore.</p>
              </div>
            `
          }
        );
      } catch (err) {
        return { status: "ERROR", message: "Failed to dispatch email. Please check configuration." };
      }

      return { status: "SUCCESS", message: "Reset code dispatched to " + storedEmail };
    }
  }

  return { status: "ERROR", message: "Mobile number not found in customer registry." };
}

function resetPassword(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const phone = String(payload.phone || "").trim();
  const resetCode = String(payload.resetCode || "").trim().toUpperCase();
  const newPassword = String(payload.newPassword || "").trim();

  if (!newPassword || newPassword.length < 4) {
    return { status: "ERROR", message: "New password must be at least 4 characters long." };
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === phone) {
      const storedCode = String(data[i][10] || "").trim().toUpperCase();
      const expires = new Date(data[i][11]);

      if (!storedCode || storedCode !== resetCode) {
        return { status: "ERROR", message: "Invalid verification code." };
      }
      if (Date.now() > expires.getTime()) {
        return { status: "ERROR", message: "Verification code has expired. Please request a new code." };
      }

      sheet.getRange(i + 1, 4).setValue(hashString(newPassword));
      sheet.getRange(i + 1, 11).setValue("");
      sheet.getRange(i + 1, 12).setValue("");

      return { status: "SUCCESS", message: "Password reset successfully. You may now log in." };
    }
  }
  return { status: "ERROR", message: "Account not found." };
}

function changePassword(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const phone = String(payload.phone || "").trim();
  const currentPassword = String(payload.currentPassword || "").trim();
  const newPassword = String(payload.newPassword || "").trim();

  if (!newPassword || newPassword.length < 4) {
    return { status: "ERROR", message: "New password must be at least 4 characters long." };
  }

  const currentHash = hashString(currentPassword);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === phone) {
      if (String(data[i][3]) !== currentHash) {
        return { status: "ERROR", message: "Current password does not match our records." };
      }

      sheet.getRange(i + 1, 4).setValue(hashString(newPassword));
      return { status: "SUCCESS", message: "Password updated successfully." };
    }
  }
  return { status: "ERROR", message: "Customer account not found." };
}

function sendChatMessage(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chat_Messages");
  const msgId = "MSG-" + Date.now();
  const custId = String(payload.customerId || "").trim();
  const custName = String(payload.customerName || "Customer").trim();
  const sender = String(payload.sender || "CUSTOMER").toUpperCase();
  const text = String(payload.messageText || "").trim();

  if (!text) return { status: "ERROR", message: "Empty message text." };

  const istTimestamp = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([msgId, istTimestamp, custId, custName, sender, text, false]);
  return { status: "SUCCESS", messageId: msgId, timestamp: istTimestamp };
}

function getChatMessages(customerId) {
  if (!customerId) return { status: "ERROR", message: "Customer ID is required." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chat_Messages");
  const data = sheet.getDataRange().getValues();
  const messages = [];
  const cleanId = String(customerId).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === cleanId) {
      messages.push({
        id: data[i][0],
        timestamp: data[i][1],
        customerId: data[i][2],
        customerName: data[i][3],
        sender: data[i][4],
        text: data[i][5],
        isRead: data[i][6]
      });
    }
  }

  return { status: "SUCCESS", messages: messages };
}

function getAllChatThreads(pin) {
  const config = getConfigData().data;
  if (String(pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chat_Messages");
  const data = sheet.getDataRange().getValues();
  const threads = {};

  for (let i = 1; i < data.length; i++) {
    const custId = String(data[i][2]).trim();
    if (!threads[custId]) {
      threads[custId] = {
        customerId: custId,
        customerName: data[i][3],
        lastMessage: data[i][5],
        lastTimestamp: data[i][1],
        sender: data[i][4],
        unreadCount: 0
      };
    }
    threads[custId].lastMessage = data[i][5];
    threads[custId].lastTimestamp = data[i][1];
    threads[custId].sender = data[i][4];
    if (data[i][4] === "CUSTOMER" && !data[i][6]) {
      threads[custId].unreadCount += 1;
    }
  }

  return { status: "SUCCESS", threads: Object.values(threads).sort((a, b) => new Date(b.lastTimestamp) - new Date(a.lastTimestamp)) };
}

function recordHeartbeat(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Active_Sessions");
  const custId = String(payload.customerId || "").trim();
  if (!custId) return { status: "ERROR", message: "Missing Customer ID." };

  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const nowIst = Utilities.formatDate(now, TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === custId) {
      sheet.getRange(i + 1, 4).setValue(nowIst);
      found = true;
      break;
    }
  }

  if (!found) sheet.appendRow([custId, payload.fullName || "Guest", payload.phone || "", nowIst]);
  return { status: "SUCCESS", timestamp: nowIst };
}

function getOnlineUsers(pin) {
  const config = getConfigData().data;
  if (String(pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Active_Sessions");
  const data = sheet.getDataRange().getValues();
  const nowMs = Date.now();
  const activeList = [];

  for (let i = 1; i < data.length; i++) {
    const lastActive = new Date(data[i][3]).getTime();
    if (nowMs - lastActive <= 60000) {
      activeList.push({ customerId: data[i][0], customerName: data[i][1], phone: data[i][2], lastActive: data[i][3] });
    }
  }

  return { status: "SUCCESS", count: activeList.length, activeUsers: activeList };
}

function adminAddMenuItem(payload) {
  const config = getConfigData().data;
  if (String(payload.pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const itemId = "ITEM-" + Math.floor(100 + Math.random() * 900);

  let imageUrl = payload.imageUrl || "https://images.unsplash.com/photo-1546833999-b9f581a1996d?w=600";
  if (payload.base64Image && payload.base64Image.length > 50) {
    try {
      const folder = getOrCreateFolder(MENU_IMAGES_FOLDER_NAME);
      const contentType = payload.base64Image.substring(payload.base64Image.indexOf(":") + 1, payload.base64Image.indexOf(";"));
      const rawBase64 = payload.base64Image.substring(payload.base64Image.indexOf(",") + 1);
      const blob = Utilities.newBlob(Utilities.base64Decode(rawBase64), contentType, "Menu_" + itemId + ".png");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = file.getUrl();
    } catch (e) {}
  }

  sheet.appendRow([
    itemId, payload.category, payload.name, payload.description,
    payload.milletType, Number(payload.price) || 0, imageUrl, true,
    Number(payload.prepTimeMins) || 15, payload.nutrition || "Balanced Fiber & Protein"
  ]);

  return { status: "SUCCESS", message: "Meal added to active menu.", itemId: itemId };
}

function adminDeleteMenuItem(payload) {
  const config = getConfigData().data;
  if (String(payload.pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.itemId).trim()) {
      sheet.deleteRow(i + 1);
      return { status: "SUCCESS", message: "Item removed from menu permanently." };
    }
  }
  return { status: "ERROR", message: "Item not found." };
}

function adminToggleMenuStock(payload) {
  const config = getConfigData().data;
  if (String(payload.pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.itemId).trim()) {
      sheet.getRange(i + 1, 8).setValue(payload.available === true);
      return { status: "SUCCESS", message: "Stock status updated." };
    }
  }
  return { status: "ERROR", message: "Item not found." };
}

function adminUpdateMenuItem(payload) {
  const config = getConfigData().data;
  if (String(payload.pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.itemId).trim()) {
      if (payload.price) sheet.getRange(i + 1, 6).setValue(Number(payload.price) || 0);
      return { status: "SUCCESS", message: "Price updated." };
    }
  }
  return { status: "ERROR", message: "Item not found." };
}

function adminUpdateConfig(payload) {
  const config = getConfigData().data;
  if (String(payload.pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.key).trim()) {
      sheet.getRange(i + 1, 2).setValue(String(payload.value));
      return { status: "SUCCESS", message: "Config key " + payload.key + " set to " + payload.value };
    }
  }
  sheet.appendRow([payload.key, payload.value, "Custom Configuration"]);
  return { status: "SUCCESS", message: "Config saved." };
}

function adminUpdateAnnouncement(payload) {
  const config = getConfigData().data;
  if (String(payload.pin) !== String(config.ADMIN_PIN)) return { status: "ERROR", message: "Unauthorized PIN." };
  return adminUpdateConfig({
    pin: payload.pin,
    key: "ANNOUNCEMENT_BANNER",
    value: String(payload.bannerText || "").trim()
  });
}
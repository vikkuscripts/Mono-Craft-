function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'update') {
      return HtmlService.createTemplateFromFile('update').evaluate()
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setTitle('Update Payment');
  }

  var refCode = e && e.parameter && e.parameter.referenceCode;

  if (refCode) {
    var template = HtmlService.createTemplateFromFile('edit');
    template.referenceCode = refCode;  
    return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createHtmlOutputFromFile('index')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('Enquiry Form')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function addPaymentToOrder(refCode, newPaymentAmount) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Sales Orders");
    if (!sh) throw new Error("Sheet 'Sales Orders' not found");

    const data = sh.getDataRange().getValues();
    if (data.length < 2) throw new Error("No data found");

    const headers = data[0].map(h => String(h).trim());
    
    // FIND COLUMNS
    const refColIdx = headers.indexOf("RefCode");
    const advColIdx = headers.indexOf("AdvanceAmount"); 
    const pdfUrlIdx = headers.indexOf("Invoice URL");
    const fileIdIdx = headers.indexOf("Invoice File ID");
    
    // Fallback if headers are slightly different (Case sensitivity or spaces)
    const findCol = (name) => headers.findIndex(h => h.toLowerCase().replace(/[^a-z0-9]/g,"") === name.toLowerCase().replace(/[^a-z0-9]/g,""));
    
    const refCol = refColIdx > -1 ? refColIdx : findCol("refcode");
    const advCol = advColIdx > -1 ? advColIdx : findCol("advanceamount");
    const pdfCol = pdfUrlIdx > -1 ? pdfUrlIdx : findCol("invoiceurl");
    const fidCol = fileIdIdx > -1 ? fileIdIdx : findCol("invoicefileid");

    if (refCol === -1) throw new Error("RefCode column not found");
    if (advCol === -1) throw new Error("AdvanceAmount column not found");

    // Search for rows
    const targets = [];
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][refCol]).trim() === String(refCode).trim()) {
            targets.push(i + 1); // 1-based row index
        }
    }

    if (targets.length === 0) throw new Error("Order not found: " + refCode);

    // GET CURRENT DATA (from first row of the order)
    const firstRowIdx = targets[0] - 1; // 0-based index for data array
    const currentAdvance = Number(data[firstRowIdx][advCol]) || 0;
    const grandTotalVal  = Number(data[firstRowIdx][findCol("GrandTotal") > -1 ? findCol("GrandTotal") : findCol("footergrandtotal")]) || 0;

    const updatedAdvance = currentAdvance + Number(newPaymentAmount);
    
    // Validation? (Optional: Prevent paying more than total? User might want to allow it)

    // UPDATE SHEET FIRST (So PDF generation if it reads from sheet - though we pass payload - is consistent? Actually we generate from payload)
    // We update all product rows with the new Advance Amount
    targets.forEach(r => {
        sh.getRange(r, advCol + 1).setValue(updatedAdvance);
    });

    // NOW REGENERATE PDF
    // We need to reconstruct payload. 
    // We reuse getOrdersByRefCode logic but locally or just call it if available?
    // getOrdersByRefCode returns structured object. We need flat payload for generateInvoicePdf_
    
    // Let's manually map the data from the first row + products
    const rowObj = {};
    headers.forEach((h, k) => rowObj[h] = data[firstRowIdx][k]);
    
    // Update the advance amount in our object
    rowObj["AdvanceAmount"] = updatedAdvance;
    
    // Map to payload schema expected by generateInvoicePdf_
    // (We use a helper here to avoid duplication if possible, or just inline)
    const pick = (field, ...keys) => {
        for(let k of keys) {
            if(rowObj[k] !== undefined) return rowObj[k];
        }
        // fuzzy search
        const normInfo = keys[0].toLowerCase().replace(/[^a-z0-9]/g,"");
         const found = headers.find(h => h.toLowerCase().replace(/[^a-z0-9]/g,"") === normInfo);
         if(found) return rowObj[found];
         return "";
    };

    const payload = {
        orderDate: pick("orderDate", "OrderDate"),
        grandTotal: pick("grandTotal", "GrandTotal", "Footer Grand Total"), // Use Footer Grand Total for invoice
        footerGrandTotal: pick("grandTotal", "GrandTotal", "Footer Grand Total"),
        advanceAmount: updatedAdvance,
        customerName: pick("customerName", "CustomerName", "Person Name"),
        address: pick("address", "BillingAddress", "Billing Address"),
        gstNumber: pick("gstNumber", "GSTNumber"),
        mobile: pick("mobile", "Mobile"),
        delCompanyName: pick("delCompanyName", "DelCompany", "DelCompanyName"),
        delAddress: pick("delAddress", "DelAddress", "DelAddress"),
        delGstNumber: pick("delGstNumber", "DelGST", "DelGSTNumber"),
        delMobile: pick("delMobile", "DelMobile"),
        delContactPerson: pick("delContactPerson", "DelPerson", "DelContactPerson"),
        user: pick("user", "User"),
        remarks: pick("remarks", "Remarks"),
        totalWithoutGst: pick("totalWithoutGst", "Total Without GST"),
        totalCgst: pick("totalCgst", "Total CGST"),
        totalSgst: pick("totalSgst", "Total SGST"),
        totalIgst: pick("totalIgst", "Total IGST"),
        products: []
    };
    
    // Collect products
    targets.forEach(r => {
        const d = data[r-1];
        const pObj = {};
        headers.forEach((h, k) => pObj[h] = d[k]);
        
        // Product payload mapping
        payload.products.push({
            productCode: pObj["ProductCode"] || pObj["Product Code"],
            productName: pObj["ProductName"] || pObj["Product Name"],
            description: pObj["Description"],
            size: pObj["Size"],
            colour: pObj["Colour"] || pObj["Color"],
            unitValue: pObj["UnitValue"] || pObj["Unit Value"],
            totalQty: pObj["Qty"] || pObj["Quantity"] || pObj["TotalQty"],
            gstPercent: pObj["GST%"] || pObj["GSTPercent"],
            totalPrice: pObj["TotalPrice"] || pObj["Total Price"],
            image: pObj["Image"] || pObj["ProductImage"] || ""
        });
    });

    // Prepare images for PDF
    payload.products = payload.products.map(p => ({
      ...p,
      imageDataUri: fetchImageAsDataUri(p.image || "")
    }));


    const invoiceInfo = generateInvoicePdf_(payload, refCode);

    // UPDATE PDF URL IN SHEET
    if (pdfCol > -1) {
        const url = invoiceInfo.pdfUrl || "";
        targets.forEach(r => {
             sh.getRange(r, pdfCol + 1).setValue(url);
        });
    }
    
    // Update File ID if exists
    if (fidCol > -1 && invoiceInfo.fileId) {
         targets.forEach(r => {
             sh.getRange(r, fidCol + 1).setValue(invoiceInfo.fileId);
        });
    }

    return {
        success: true,
        message: "Payment updated and PDF regenerated!",
        newBalance: (grandTotalVal - updatedAdvance),
        totalPaid: updatedAdvance,
        pdfUrl: invoiceInfo.pdfUrl
    };

  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function getGstType() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('GST Type');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(r => r[0]).filter(String);
}
function getFy() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('FY');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(r => r[0]).filter(String);
}

// function getOrderCategory() {
//   const sheet = SpreadsheetApp.getActive().getSheetByName('Order Category');
//   if (!sheet) return [];
//   const data = sheet.getDataRange().getValues();
//   return data.slice(1).map(r => r[0]).filter(String);
// }

function getClientType() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Client Type');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(r => r[0]).filter(String);
}
function getHeadsList() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Head');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift(); // remove header row
  return data.map(r => r[0]).filter(String);
}
function getUsersByHead(selectedHead) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Head');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  const row = data.find(r => r[0] === selectedHead);
  if (!row) return [];
  return row.slice(1).filter(String);
}

function searchProductByName(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Stock Sheet");
  const values = sh.getDataRange().getValues();
  const headers = values.shift();

  const idx = {
    code: headers.indexOf("Code"),
    productName: headers.indexOf("Product Name"),
    size: headers.indexOf("Size"),
    color: headers.indexOf("Color"),
    opening: headers.indexOf("Opening Stock"),
    inward: headers.indexOf("Inward"),
    outward: headers.indexOf("Outward"),
    closing: headers.indexOf("Closing Stock"),
    description: headers.indexOf("Description"),
    price: headers.indexOf("Price"),
    productCategory: headers.indexOf("Product Category"),
    hsnCode: headers.indexOf("HSN Code"),
    image: headers.indexOf("Image") // <-- Added image index
  };

  return values
    .filter(r => (r[idx.productName] + "").toLowerCase().includes(name.toLowerCase()))
    .map(r => ({
      code: r[idx.code],
      productName: r[idx.productName],
      size: r[idx.size],
      color: r[idx.color],
      openingStock: r[idx.opening],
      inward: r[idx.inward],
      outward: r[idx.outward],
      closingStock: r[idx.closing],
      description: r[idx.description],
      price: r[idx.price],
      productCategory: r[idx.productCategory],
      hsnCode: r[idx.hsnCode],
      image: r[idx.image] || "" // <-- Return image value safely
    }));
}


function getAllProducts() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Stock Sheet");
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const headers = values[0].map(String);
  const rows = values.slice(1);

  const idx = {
    code: headers.indexOf("Code"),
    productName: headers.indexOf("Product Name"),
    size: headers.indexOf("Size"),
    color: headers.indexOf("Color"),
    opening: headers.indexOf("Opening Stock"),
    inward: headers.indexOf("Inward"),
    outward: headers.indexOf("Outward"),
    closing: headers.indexOf("Closing Stock"),
    description: headers.indexOf("Description"),
    price: headers.indexOf("Price"),
    productCategory: headers.indexOf("Product Category"),
    hsnCode: headers.indexOf("HSN Code"),
    image: headers.indexOf("Image") // <-- Newly added column
  };

  return rows.map(r => ({
    code: idx.code >= 0 ? r[idx.code] : "",
    productName: idx.productName >= 0 ? r[idx.productName] : "",
    size: idx.size >= 0 ? r[idx.size] : "",
    color: idx.color >= 0 ? r[idx.color] : "",
    openingStock: idx.opening >= 0 ? r[idx.opening] : "",
    inward: idx.inward >= 0 ? r[idx.inward] : "",
    outward: idx.outward >= 0 ? r[idx.outward] : "",
    closingStock: idx.closing >= 0 ? r[idx.closing] : "",
    description: idx.description >= 0 ? r[idx.description] : "",
    price: idx.price >= 0 ? r[idx.price] : "",
    productCategory: idx.productCategory >= 0 ? r[idx.productCategory] : "",
    hsnCode: idx.hsnCode >= 0 ? r[idx.hsnCode] : "",
    image: idx.image >= 0 ? r[idx.image] : ""   // <-- Return image link
  }));
}

function formatDMY(dateStr) {
  if (!dateStr) return "";
  const parts = dateStr.split("-"); // yyyy-mm-dd
  return parts[2] + "/" + parts[1] + "/" + parts[0]; // dd/mm/yyyy
}

// ------------- CONFIG -------------
// ------------- CONFIG -------------
const DRIVE_FOLDER_ID = "1Cz6lCyTQa3wZOn9D0MLZzErWw7s8QpT0";
const SHEET_NAME      = "Sales Orders";
const REF_COL         = 2;           // col B
const DATA_START_ROW  = 2;
const REF_PREFIX      = "R";
const REF_DEFAULT_START = 10001;
const REF_PAD_TO      = 0;

// images sheet name (for logo + signature)
const IMAGES_SHEET_NAME = "image";
// ----------------------------------

/**
 * Helpers
 */
function formatDMY(iso) {
  if (!iso) return "";
  try {
    const parts = String(iso).split("-");
    if (parts.length !== 3) return iso;
    const [y, m, d] = parts;
    return `${d}/${m}/${y}`;
  } catch (e) {
    return iso;
  }
}

function formatInvoiceDate(iso) {
  if (!iso) return "";
  try {
    const parts = String(iso).split("-");
    if (parts.length !== 3) return iso;
    const [y, m, d] = parts;
    const monthNames = [
      "JAN","FEB","MAR","APR","MAY","JUN",
      "JUL","AUG","SEP","OCT","NOV","DEC"
    ];
    const monthIndex = parseInt(m, 10) - 1;
    const mm = monthNames[monthIndex] || m;
    return `${parseInt(d,10)} ${mm} ${y}`;
  } catch (e) {
    return iso;
  }
}

/**
 * Extract Drive fileId from various google drive URLs
 */
function extractDriveFileId(url) {
  if (!url) return null;
  url = String(url).trim();
  // common patterns:
  // https://drive.google.com/file/d/<id>/view...
  // https://drive.google.com/open?id=<id>
  // https://drive.google.com/uc?id=<id>&export=download
  const patterns = [
    /\/d\/([a-zA-Z0-9_-]+)(?:[\/?]|$)/,
    /id=([a-zA-Z0-9_-]+)(?:&|$)/,
    /\/file\/u\/\d+\/d\/([a-zA-Z0-9_-]+)(?:[\/?]|$)/
  ];
  for (let i = 0; i < patterns.length; i++) {
    const m = url.match(patterns[i]);
    if (m && m[1]) return m[1];
  }
  // maybe it's already an id
  if (/^[a-zA-Z0-9_-]{10,}$/i.test(url)) return url;
  return null;
}

/**
 * Convert a Blob to a data URI string (base64)
 */
function blobToDataUri(blob) {
  if (!blob) return "";
  try {
    const bytes = blob.getBytes();
    const base64 = Utilities.base64Encode(bytes);
    const mime = blob.getContentType() || "application/octet-stream";
    return "data:" + mime + ";base64," + base64;
  } catch (e) {
    return "";
  }
}

/**
 * Try to fetch image as a Blob from either:
 * - Google Drive file id (using DriveApp)
 * - Remote https/http url (using UrlFetchApp)
 * Returns a data URI (or empty string on failure)
 */
function fetchImageAsDataUri(rawUrlOrId) {
  if (!rawUrlOrId) return "";
  const s = String(rawUrlOrId).trim();
  // first try extracting a drive file id
  const fileId = extractDriveFileId(s);
  if (fileId) {
    try {
      const file = DriveApp.getFileById(fileId);
      const blob = file.getBlob();
      return blobToDataUri(blob);
    } catch (e) {
      // Not accessible via DriveApp (permissions) — fallthrough to try fetch below
      // Logger.log("Drive fetch failed for id " + fileId + ": " + e);
    }
  }

  // If it looks like an http(s) url, try UrlFetchApp
  if (/^https?:\/\//i.test(s)) {
    try {
      const response = UrlFetchApp.fetch(s, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true, timeout: 30000 });
      if (response.getResponseCode && response.getResponseCode() === 200) {
        const blob = response.getBlob();
        return blobToDataUri(blob);
      } else {
        // Logger.log("UrlFetchApp response not 200: " + (response.getResponseCode ? response.getResponseCode() : "unknown"));
      }
    } catch (e) {
      // Logger.log("UrlFetchApp.fetch failed for " + s + ": " + e);
    }
  }

  // last resort: treat input as raw base64 data uri already
  if (s.indexOf("data:") === 0) {
    return s;
  }

  return "";
}

/**
 * Read logo + signature from "image" sheet and return data URIs.
 * Expect headers in row1: "Logo URL", "Signature URL" (case-insensitive)
 */
function getInvoiceImageUrls_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(IMAGES_SHEET_NAME);
  if (!sh) {
    return { logoDataUri: "", signatureDataUri: "" };
  }

  const lastCol = sh.getLastColumn();
  if (lastCol === 0) {
    return { logoDataUri: "", signatureDataUri: "" };
  }

  // read first two rows: header + first data row
  const data = sh.getRange(1, 1, 2, lastCol).getValues();
  const headers = data[0] || [];
  const row = data[1] || [];

  const map = {};
  headers.forEach((h, i) => {
    if (h) map[String(h).trim().toLowerCase()] = row[i];
  });

  // Accept multiple header name variants
  const logoRaw = map["logo url"] || map["logo"] || map["logo_url"] || "";
  const signRaw = map["signature url"] || map["signature"] || map["signature_url"] || map["sign"] || "";

  // Convert to data URIs
  const logoDataUri = fetchImageAsDataUri(logoRaw);
  const signatureDataUri = fetchImageAsDataUri(signRaw);

  return {
    logoDataUri: String(logoDataUri || "").trim(),
    signatureDataUri: String(signatureDataUri || "").trim()
  };
}

/**
 * Convert number to words in Indian style (up to crores)
 */
function numberToIndianWords_(num) {
  num = Number(num);
  if (isNaN(num)) return "";

  if (num === 0) return "Zero";

  const a = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen",
    "Seventeen","Eighteen","Nineteen"
  ];
  const b = [
    "", "", "Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"
  ];

  function inWords(n) {
    let str = "";

    if (n > 19) {
      str += b[Math.floor(n / 10)] + (n % 10 ? " " + a[n % 10] : "");
    } else {
      str += a[n];
    }

    return str;
  }

  let s = "";
  const crore     = Math.floor(num / 10000000);
  const lakh      = Math.floor((num / 100000) % 100);
  const thousand  = Math.floor((num / 1000) % 100);
  const hundred   = Math.floor((num / 100) % 10);
  const rest      = Math.floor(num % 100);

  if (crore > 0) {
    s += inWords(crore) + " Crore ";
  }
  if (lakh > 0) {
    s += inWords(lakh) + " Lakh ";
  }
  if (thousand > 0) {
    s += inWords(thousand) + " Thousand ";
  }
  if (hundred > 0) {
    s += inWords(hundred) + " Hundred ";
  }
  if (rest > 0) {
    if (s !== "") s += "And ";
    s += inWords(rest) + " ";
  }

  return s.trim();
}

function amountToWordsINR_(num) {
  num = Number(num);
  if (isNaN(num)) return "";

  const rupees = Math.floor(num);
  const paise  = Math.round((num - rupees) * 100);

  let result = numberToIndianWords_(rupees) + " Rupees";
  if (paise > 0) {
    result += " And " + numberToIndianWords_(paise) + " Paise";
  }
  result += " Only";

  return result.toUpperCase();
}

/**
 * Build invoice HTML using payload + products + refCode + logo/signature data URIs
 */
function buildInvoiceHtml_(payload, refCode, logoDataUri, signatureDataUri) {
  const orderDateText  = formatInvoiceDate(payload.orderDate);
  const totalWithGst   = Number(payload.grandTotal || 0);
  const totalPaid      = Number(payload.advanceAmount || 0);
  const totalBalance   = totalWithGst - totalPaid;
  const amountInWords  = amountToWordsINR_(totalWithGst);

  // Build product rows
  let rowsHtml = "";
  (payload.products || []).forEach((p, idx) => {
    const sNo = idx + 1;
    rowsHtml += `
      <tr>
        <td class="text-center">${sNo}</td>
        <td class="text-center">
  ${p.imageDataUri ?
    `<img src="${p.imageDataUri}" style="width:45px;height:45px;object-fit:contain;" />`
    : "-"
  }
</td>

        <td class="text-center">${p.productCode || ""}</td>
        <td>${p.productName || ""}</td>
        <td>${p.description || ""}</td>
        <td class="text-center">${p.size || ""}</td>
        <td class="text-center">${p.colour || ""}</td>
        <td class="text-right">${p.unitValue || ""}</td>
        <td class="text-center">${p.totalQty || ""}</td>
        <td class="text-center">${p.gstPercent || ""}</td>
        <td class="text-right">${p.totalPrice || ""}</td>
      </tr>
    `;
  });

  // Company details
  const companyName = "FIBRE CRAFTS CONTINENTAL PRIVATE LIMITED";
  const factoryAdd  = "82, Magazine Chowk, Bhosari Alandi Road, Pune, Maharashtra - 412105";
  const officeAdd   = "HB3/2, 04 Ajmera Complex, Pimpri, Pune - 411018";
  const phoneLine   = "9975202122, 7558607221, Punit-8669202122, Sumit-8149021125,<br> Kapil-8411961551, Pooja-8446381125, Mohini-8149871125, Seema-7264821255";
  const gstin       = "27ABNCS3037L1ZV";

  const custName    = payload.customerName || "";
  const custAddr    = payload.address || "";
  const custGst     = payload.gstNumber || "";
  const custMobile  = payload.mobile || "";
  const custContact = payload.customerName || "";

  const consName    = payload.delCompanyName || "";
  const consAddr    = payload.delAddress || "";
  const consGst     = payload.delGstNumber || "";
  const consMobile  = payload.delMobile || "";
  const consPerson  = payload.delContactPerson || "";

  // Bank details
  const bankName    = "KOTAK MAHINDRA BANK";
  const bankBranch  = "PIMPRI CHINCHWAD, PUNE";
  const bankAccNo   = "6620331155";
  const bankIfsc    = "KKBK0000725";

  const totalWithoutGst  = Number(payload.totalWithoutGst || 0).toFixed(2);
  const totalCgst        = Number(payload.totalCgst || 0).toFixed(2);
  const totalSgst        = Number(payload.totalSgst || 0).toFixed(2);
  const totalIgst        = Number(payload.totalIgst || 0).toFixed(2);
  const footerGrandTotal = Number(payload.footerGrandTotal || payload.grandTotal || 0).toFixed(2);

  const totalPaidText = totalPaid.toFixed(2);
  const totalBalText  = totalBalance.toFixed(2);

  const salesRep = payload.user || "";

  // Minor styling: ensure embedded images are max-width and look good in PDF
  const logoImgHtml = logoDataUri ? `<img src="${logoDataUri}" style="max-width:95px; max-height:55px;" />` : "";
  const sigImgHtml  = signatureDataUri ? `<img src="${signatureDataUri}" style="max-height:60px;" />` : "";


const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8" />
  <style>
    @media print {
      body {
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
    }    * { box-sizing: border-box; }

    body {
      font-family: Arial, sans-serif;
      font-size: 10px;
      margin: 0;
      padding: 10px 18px 20px 18px;
      background-color: #ffffff;
    }

    /* Make all borders thin */
    table {
      border-collapse: collapse;
      width: 100%;
    }
    table, th, td {
      border: 0.1px solid #000;
    }

    th, td {
      padding: 3px 4px;
      vertical-align: top;
    }

    th { font-weight: bold; }

    .no-border {
      border: none !important;
    }

    .text-right  { text-align: right; }
    .text-center { text-align: center; }
    .text-left   { text-align: left; }

    .turmeric-fill {
      background-color: #FFC533 !important;
    }
    .turmeric-inline {
      background-color: #FFC533 !important;
      padding: 0 2px;
    }

    /* HEADER */
    .header-table {
      border: none;
      margin-bottom: 4px;
    }
    .header-table td {
      border: none;
    }

    .company-name {
      color: #0A1F63;
      font-weight: bold;
      font-size: 20px;
      text-transform: uppercase;
      margin-bottom: 2px;
      box-shadow:2px 1px 2px solid #000000;
    }

    .header-logo {
      width: 70%;
      text-align: right;
      vertical-align: middle;
    }

    /* CUSTOMER & CONSIGNEE BLOCK */
    .party-block {
      margin-bottom: 4px;
      font-size: 9px;
    }

    .party-block .title-row td {
      background-color: #FFC533;
      color: #000;
      font-weight: bold;
      text-align: center;
      font-size: 11px;
      padding: 6px 4px;
    }

    .party-cell {
      padding: 4px;
    }

    .party-name {
      font-weight: bold;
      margin-bottom: 3px;
      font-size: 10px;
    }

    .party-line {
      margin-bottom: 2px;
    }

    .party-label {
      font-weight: bold;
      display: inline-block;
      min-width: 80px;
    }

    .party-value {
      display: inline;
    }

    /* PRODUCT TABLE */
    .product-table thead th {
      background-color: #1877F2;
      color: #ffffff;
      text-align: center;
    }

    /* Section title bars (if any other use) */
    .section-title-bar {
      background-color: #1877F2;
      color: #ffffff;
      font-weight: bold;
      text-align: center;
      font-size: 11px;
      padding: 2px 0;
    }

    .fill-yellow {
      background-color: #FFC533 !important;
      border: 0.5px solid #000 !important;
      color: #000 !important;
    }

    /* Remarks, Bank & Signature blocks */
    .remarks-bank-sign-table {
      margin-top: 6px;
      border-collapse: collapse;
      width: 100%;
    }

    .block-box {
      padding: 4px;
      border: 0.5px solid #000;
      vertical-align: top;
      font-size: 9px;
    }

    .remarks-title {
      font-weight: bold;
    }

    .bank-title {
      font-weight: bold;
      margin-top: 4px;
      margin-bottom: 2px;
    }

    .signature-block {
      text-align: right;
      font-size: 9px;
    }

    .signature-stamp {
      margin: 5px 0 6px 0;
      min-height: 60px;
    }

    /* Amount in words */
    .amount-words {
      margin-top: 6px;
      border: 0.5px solid #000;
      padding: 4px;
      font-size: 9px;
    }

    /* Payment term */
    .payment-term {
      margin-top: 6px;
      font-size: 9px;
      border: 0.5px solid #000;
      padding: 4px;
    }

    .payment-term-title {
      font-weight: bold;
      margin-right: 4px;
    }

    /* Footer note */
    .footer-note {
      margin-top: 4px;
      font-size: 8px;
      text-align: center;
    }

    .bottom-status {
    width: 100%;
    border-collapse: collapse;
    margin-top: 10px;
  }

  .bottom-status td {
    border: none !important; /* Remove table borders */
    padding: 4px 6px;
    font-size: 11px;
  }

  .bottom-label {
    font-weight: bold;
    white-space: nowrap;
  }

  /* Small inline blank box */
  .bottom-box {
    display: inline-block;
    width: 120px;  /* adjust size if needed */
    border-bottom: 1px solid #000; /* Underline blank field */
    height: 12px;
    margin-left: 6px;
  }
    .remarks-bank-sign-table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 10px;
  }
  .remarks-bank-sign-table td {
    vertical-align: top;
    padding: 8px 10px;
    border: 1px solid #000; /* Only outer border */
  }

  /* Inner sections split equally */
  .left-section {
    display: flex;
    flex-direction: column;
    height: 100%;
  }
  .remarks-section,
  .bank-section {
    flex: 1;
    padding-bottom: 10px;
  }

  .remarks-title,
  .bank-title {
    font-weight: bold;
    text-decoration: underline;
    display: inline-block;
  }

  .bank-row {
    display: flex;
    justify-content: space-between;
    width: 100%;
    line-height: 10px;
  }

  .bank-label {
    font-weight: bold;
    width: 45%;
  }
  .bank-value {
    width: 53%;
    text-align: left;
  }

  .signature-stamp {
    margin-top: 20px;
    margin-bottom: 2px;
    margin-left:5px;
  }
/* Increase signature image size */
.signature-stamp img {
  max-height: 120px;   /* increase number for bigger signature */
  width: auto;
  object-fit: contain;
}

  </style>
</head>
<body>

  <!-- HEADER -->
  <table class="header-table">
    <tr>
      <td style="width: 80%;">
        <div class="company-name">${companyName}</div>
        <div>Fact.Add. : ${factoryAdd}</div>
        <div>Off.Add. : ${officeAdd}</div>
        <div>Phone No. : ${phoneLine}</div>
        <div>GSTIN NO. : ${gstin}</div>
      </td>
      <td class="header-logo" style="width: 40%;">
        ${logoImgHtml}
      </td>
    </tr>
  </table>

<!-- CUSTOMER & CONSIGNEE BLOCK -->
<table class="party-block">
  <!-- header row with turmeric background -->
  <tr class="title-row">
    <td>Customer Details</td>
    <td>Consignee Details</td>
  </tr>

  <!-- Content row with aligned label-value format -->
  <tr>
    <!-- Customer Block -->
    
    <td class="party-cell">
      <div class="party-name">${custName}</div>

      <div class="party-line">
        <span class="party-label">Address</span>
        <span class="colon">:</span>
        <span class="party-value">${custAddr}</span>
      </div>

      <div class="party-line">
        <span class="party-label">GST</span>
        <span class="colon">:</span>
        <span class="party-value">${custGst}</span>
      </div>

      <div class="party-line">
        <span class="party-label">Contact Person</span>
        <span class="colon">:</span>
        <span class="party-value">${custContact}</span>
      </div>

      <div class="party-line">
        <span class="party-label">Contact No.</span>
        <span class="colon">:</span>
        <span class="party-value">${custMobile}</span>
      </div>
    </td>

    <!-- Consignee Block -->
    <td class="party-cell">
      <div class="party-name">${consName}</div>

      <div class="party-line">
        <span class="party-label">Address</span>
        <span class="colon">:</span>
        <span class="party-value">${consAddr}</span>
      </div>

      <div class="party-line">
        <span class="party-label">GST</span>
        <span class="colon">:</span>
        <span class="party-value">${consGst}</span>
      </div>

      <div class="party-line">
        <span class="party-label">Contact Person</span>
        <span class="colon">:</span>
        <span class="party-value">${consPerson}</span>
      </div>

      <div class="party-line">
        <span class="party-label">Contact No.</span>
        <span class="colon">:</span>
        <span class="party-value">${consMobile}</span>
      </div>
    </td>
  </tr>
</table>


  <!-- SALES ORDER INFO -->
  <table class="sales-info" style="margin-bottom:4px; font-size:10px;">
    <tr>
      <td style="width:50%;">
        <b>Sales Order No. :</b> ${refCode}
      </td>
      <!-- turmeric background on Order Date -->
      <td class="turmeric-fill" style="width:50%;">
        <b>Sales Order date :</b> ${orderDateText}
      </td>
    </tr>
    <tr>
      <td colspan="2" style="padding:6px 8px; text-align:left;">
        <b>Sales Person :</b> ${salesRep}
      </td>
    </tr>
  </table>

  <!-- PRODUCT TABLE -->
  <table class="product-table">
    <thead>
      <tr>
        <th style="width:3%;">S.No.</th>
        <th style="width:10%;">Image</th>
        <th style="width:5%;">Product<br>Code</th>
        <th style="width:8%;">Product Name</th>
        <th>Description</th>
        <th style="width:5%;">Size</th>
        <th style="width:6%;">Color</th>
        <th style="width:8%;">Unit Value</th>
        <th style="width:3%;">QTY</th>
        <th style="width:4%;">GST%</th>
        <th style="width:8%;">Total Price</th>
      </tr>
    </thead>
    <tbody>
      ${rowsHtml}
    </tbody>
  </table>

  <!-- TOTALS -->
  <table style="margin-top:6px;">
    <tr>
      <td>Total Without GST</td>
      <td class="text-right">${totalWithoutGst}</td>
    </tr>
    <tr>
      <td>CGST Amount</td>
      <td class="text-right">${totalCgst}</td>
    </tr>
    <tr>
      <td>SGST Amount</td>
      <td class="text-right">${totalSgst}</td>
    </tr>
    <tr>
      <td>IGST Amount</td>
      <td class="text-right">${totalIgst}</td>
    </tr>
    <!-- turmeric background on Total (grand total) -->
    <tr>
      <td class="turmeric-fill"><b>Total With GST</b></td>
      <td class="turmeric-fill text-right"><b>${footerGrandTotal}</b></td>
    </tr>
    <tr>
      <td>Total Paid</td>
      <!-- turmeric background on Paid Amount -->
      <td class="turmeric-fill text-right">${totalPaidText}</td>
    </tr>
    <tr>
      <td>Total Balance</td>
      <td class="text-right">${totalBalance.toFixed(2)}</td>
    </tr>
  </table>

  <!-- AMOUNT IN WORDS -->
  <div class="amount-words">
    Total Amount in Words : ${amountInWords}
  </div>

 <table class="remarks-bank-sign-table" style="width:100%; border-collapse:collapse; margin-top:10px; font-size:9px;">
  <tr>
    <!-- LEFT: Remarks (Terms) -->
    <td style="width:55%; border:1px solid #000; vertical-align:top; padding:6px;">
       <div style="font-weight:bold; text-decoration:underline; margin-bottom:5px;">Terms & Condition :</div>
       <div style="white-space:pre-wrap; font-family:sans-serif; line-height:1.4;">${payload.remarks || ""}</div>
    </td>

    <!-- RIGHT: Bank Details + Signature -->
    <td style="width:45%; border:1px solid #000; vertical-align:top; padding:6px;">
       
       <!-- Bank Details (Top) -->
       <div style="margin-bottom:10px;">
         <div style="font-weight:bold; text-decoration:underline; margin-bottom:2px;">Bank Details</div>
         <table style="width:100%; border-collapse:collapse; font-size:9px; border:none; margin:0;">
           <tr style="border:none;">
             <td style="border:none; padding:0 2px; width:70px; font-weight:bold;">Bank Name</td>
             <td style="border:none; padding:0 2px; width:10px;">:</td>
             <td style="border:none; padding:0 2px;">${bankName}</td>
           </tr>
           <tr style="border:none;">
             <td style="border:none; padding:0 2px; font-weight:bold;">Branch</td>
             <td style="border:none; padding:0 2px;">:</td>
             <td style="border:none; padding:0 2px;">${bankBranch}</td>
           </tr>
           <tr style="border:none;">
             <td style="border:none; padding:0 2px; font-weight:bold;">Account No.</td>
             <td style="border:none; padding:0 2px;">:</td>
             <td style="border:none; padding:0 2px;">${bankAccNo}</td>
           </tr>
           <tr style="border:none;">
             <td style="border:none; padding:0 2px; font-weight:bold;">IFSC</td>
             <td style="border:none; padding:0 2px;">:</td>
             <td style="border:none; padding:0 2px;">${bankIfsc}</td>
           </tr>
         </table>
       </div>

       <!-- Signature (Bottom) -->
       <div style="text-align:right; margin-top:20px;">
          For: <b>${companyName}</b>
          <div class="signature-stamp" style="margin:5px 0 5px auto; min-height:50px;">${sigImgHtml}</div>
          <div style="font-weight:bold;">Authorised Signatory</div>
       </div>
    </td>
  </tr>
</table>


  <!-- FOOTER NOTE -->
  <div class="footer-note">
    This is a computer generated proforma. No signature is required.
  </div>

  <!-- BOTTOM STATUS -->
 <table class="bottom-status">
  <tr>
    <td class="bottom-label">Made By:</td>
    <td><span class="bottom-box"></span></td>

    <td class="bottom-label">Checked By:</td>
    <td><span class="bottom-box"></span></td>

    <td class="bottom-label">Approved By:</td>
    <td><span class="bottom-box"></span></td>
  </tr>

  <tr>
    <td class="bottom-label">MFG Days:</td>
    <td><span class="bottom-box" style="width:80px;"></span></td>

    <td class="bottom-label">Dispatch Date:</td>
    <td><span class="bottom-box" style="width:80px;"></span></td>

    <td></td><td></td> <!-- empty to balance width -->
  </tr>
</table>
</body>
</html>
`;


return html;
}


/**
 * Create invoice PDF from payload & refCode, returns { pdfUrl, fileId }
 */
function generateInvoicePdf_(payload, refCode) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  // get image data URIs
  const images = getInvoiceImageUrls_();
  const logoDataUri = images.logoDataUri || "";
  const signatureDataUri = images.signatureDataUri || "";

  const html = buildInvoiceHtml_(payload, refCode, logoDataUri, signatureDataUri);

  const output = HtmlService
    .createHtmlOutput(html)
    .setWidth(800)
    .setHeight(1123); // approx A4

  const blob = output
    .getBlob()
    .getAs('application/pdf')
    .setName("Invoice_" + refCode + ".pdf");

  const file = folder.createFile(blob);

  return {
    pdfUrl: file.getUrl(),
   
  };
}

/**
 * Main save function (unchanged logic, but ensures invoiceInfo.fileId returned)
 */
function saveSalesOrderWithFiles(payload, files, opts) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000); // avoid ref collisions

    // ===== Spreadsheet =====
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error("Sheet '" + SHEET_NAME + "' not found.");

    if (!payload || !Array.isArray(payload.products) || payload.products.length === 0) {
      throw new Error("No products found in payload.");
    }

    // ===== Timestamp =====
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm:ss"
    );

    // ===== Generate new Reference Code =====
    const lastRow = Math.max(sh.getLastRow(), DATA_START_ROW - 1);
    let nextRefNumber = REF_DEFAULT_START;

    if (lastRow >= DATA_START_ROW) {
      const numRows   = lastRow - (DATA_START_ROW - 1);
      const refValues = sh.getRange(DATA_START_ROW, REF_COL, numRows, 1).getValues();

      for (let i = refValues.length - 1; i >= 0; i--) {
        let v = refValues[i][0];
        if (!v) continue;
        v = String(v).trim();

        const m = v.match(/^\s*R?0*(\d+)\s*$/i);
        if (m && m[1]) {
          const parsed = parseInt(m[1], 10);
          if (!isNaN(parsed)) {
            nextRefNumber = parsed + 1;
            break;
          }
        }
      }
    }

    let numeric = String(nextRefNumber);
    if (REF_PAD_TO && REF_PAD_TO > 0) numeric = numeric.padStart(REF_PAD_TO, "0");
  const refCode = REF_PREFIX + numeric;

    // ===== File Upload Handling =====
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    
    // Initialize with existing URLs from payload (hidden inputs)
    let customerDocURL  = payload.existingCustomerDocUrl || "";
    let agreementDocURL = payload.existingAgreementCopyUrl || "";

    if (files && files.length) {
      files.forEach(f => {
        try {
          const blob = Utilities.newBlob(
            Utilities.base64Decode(f.data),
            f.mimeType || "application/octet-stream",
            f.fileName
          );
          const saved = folder.createFile(blob);
          const url = saved.getUrl();

          if (f.fieldName === "customerDocs")      customerDocURL  = url;
          else if (f.fieldName === "agreementCopy") agreementDocURL = url;

        } catch (e) {
          Logger.log("File upload failed: " + (f && f.fileName) + " → " + e);
        }
      });
    }
// ===== Prepare product images for invoice (convert to dataURI) =====
payload.products = payload.products.map(p => {
  return {
    ...p,
    imageDataUri: fetchImageAsDataUri(p.productImage || p.image || "")
  };
});

// ===== Build Rows for Sales Orders Sheet (NO IMAGE SAVED) =====
const rows = payload.products.map(p => [
  timestamp,                
  refCode,                  
  formatDMY(payload.orderDate),
  payload.fy,
  payload.gstType,
  payload.mobile,
  payload.altMobile,
  payload.email,
  payload.customerName,
  payload.companyName,
  payload.address,
  payload.gstNumber,
  // formatDMY(payload.closerDate),
  payload.heads,
  payload.user,

  payload.deliverySameAsOrder ? "Yes" : "No",
  payload.delMobile,
  payload.delGstNumber,
  payload.delCompanyName,
  payload.delContactPerson,
  payload.delEmail,
  payload.delAddress,

  p.productCode,
  p.productName,
  p.description,
  p.size,
  p.colour,
  p.hsn,
  p.totalQty,
  p.unitValue,
  p.stock,
  p.subTotal,
  p.gstPercent,
  p.gstAmount,
  p.totalPrice,

  payload.grandTotal,
  payload.poNumber,
  payload.poDate,
  payload.advanceAmount,
  payload.clientType,
  "", // payload.paymentTerms removed
  payload.remarks,

  payload.totalWithoutGst,
  payload.totalCgst,
  payload.totalSgst,
  payload.totalIgst,
  payload.footerGrandTotal,

  customerDocURL,
  agreementDocURL
]);


    // ===== Write to Sheet in one batch =====
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    // ===== Generate Invoice PDF (from same payload) =====
    const invoiceInfo = generateInvoicePdf_(payload, refCode);

    // ===== Save Invoice URL + File ID into sheet =====
    try {
      const invoiceData = [];
      for (let i = 0; i < rows.length; i++) {
        invoiceData.push([invoiceInfo.pdfUrl || "", invoiceInfo.fileId || ""]);
      }
      // col 53 = Invoice PDF URL, col 54 = Invoice File ID
      sh.getRange(startRow, 49, rows.length, 2).setValues(invoiceData);
    } catch (e) {
      Logger.log("Failed to write invoice URL to sheet: " + e);
    }

    return {
      success: true,
      message: "Sales Order & Invoice saved successfully!",
      referenceCode: refCode,
      rowsAdded: rows.length,
      invoiceUrl: invoiceInfo.pdfUrl,
     
    };

  } catch (err) {
    return {
      success: false,
      message: err && err.message ? err.message : String(err)
    };

  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


// 
// Edit Process code 
// 




function getNextRefCodeFromSheet_(sh) {
  const lastRow = Math.max(sh.getLastRow(), DATA_START_ROW - 1);
  let maxNum = REF_DEFAULT_START - 1;

  if (lastRow >= DATA_START_ROW) {
    const numRows = lastRow - (DATA_START_ROW - 1);
    const refValues = sh
      .getRange(DATA_START_ROW, REF_COL, numRows, 1)
      .getValues();

    for (let i = 0; i < refValues.length; i++) {
      const v = refValues[i][0];
      if (!v) continue;

      const m = String(v).match(/(\d+)/);
      if (m) {
        const n = parseInt(m[1], 10);
        if (!isNaN(n) && n > maxNum) maxNum = n;
      }
    }
  }

  const nextNum = maxNum + 1;
  return REF_PREFIX + String(nextNum).padStart(REF_PAD_TO || 0, "0");
}




function getOrdersByRefCode(refCode) {
  try {

    if (!refCode && refCode !== 0) return [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Sales Orders");
    if (!sheet) return { error: "Sheet 'Sales Orders' not found" };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    // ---------- Normalize Headers ----------
    const headers = data[0].map(h =>
      String(h || "")
        .replace(/^\uFEFF/, "")
        .trim()
    );

    const refCol = headers.indexOf("RefCode");
    if (refCol === -1) {
      return { error: "RefCode column not found", headers };
    }

    const normalizedRef = String(refCode).trim();

    let orderRows = [];

    // ---------- Collect All Rows ----------
    for (let i = 1; i < data.length; i++) {

      if (String(data[i][refCol]).trim() !== normalizedRef) continue;

      let rowObj = {};

      headers.forEach((header, index) => {
        let value = data[i][index];

        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }

        // Keep numbers as numbers
        if (value === "" || value === null) value = "";

        rowObj[header] = value;
      });

      orderRows.push(rowObj);
    }

    if (orderRows.length === 0) return [];

    // ---------- Build Structured Order ----------
    const firstRow = orderRows[0];

    // Collect Items
    const items = orderRows.map(r => ({
      ProductCode: r["ProductCode"],
      ProductName: r["ProductName"],
      Description: r["Description"],
      Size: r["Size"],
      Colour: r["Colour"],
      HSN: r["HSN"],
      Qty: Number(r["Qty"] || 0),
      UnitValue: Number(r["UnitValue"] || 0),
      SubTotal: Number(r["SubTotal"] || 0),
      GSTPercent: Number(r["GST%"] || 0),
      GSTAmount: Number(r["GSTAmount"] || 0),
      TotalPrice: Number(r["TotalPrice"] || 0)
    }));

    // ---------- Totals ----------
    const totals = {
      GrandTotal: Number(firstRow["GrandTotal"] || 0),
      AdvanceAmount: Number(firstRow["AdvanceAmount"] || 0),
      Balance: Number(firstRow["Footer Grand Total"] || 0)
    };

    // ---------- Final Order Object ----------
    const order = {
      header: firstRow,   // customer + order details
      items: items,
      totals: totals
    };

    return order;

  } catch (err) {
    Logger.log(err.stack);
    return { error: String(err) };
  }
}

function test_getOrdersByRefCode() {

  const refCode = "R10090"; // 🔹 Change this to test any reference

  const order = getOrdersByRefCode(refCode);

  Logger.log("===== FULL ORDER OBJECT =====");
  Logger.log(JSON.stringify(order, null, 2));

  if (!order || order.length === 0) {
    Logger.log("No order found");
    return;
  }

  Logger.log("===== HEADER DATA =====");
  Logger.log(JSON.stringify(order.header, null, 2));

  Logger.log("===== ITEMS =====");
  order.items.forEach((item, index) => {
    Logger.log("Item " + (index + 1));
    Logger.log(JSON.stringify(item, null, 2));
  });

  Logger.log("===== TOTALS =====");
  Logger.log(JSON.stringify(order.totals, null, 2));
}

function getItemImageByCode(itemCode) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Stock Sheet"); // 👈 Change name if needed

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    if (String(data[i][0]).trim() === String(itemCode).trim()) {

      return data[i][12] || ""; // Image column
    }
  }

  return "";
}
function test_getItemImageByCode() {

  const code = "F2/F";

  const image = getItemImageByCode(code);

  Logger.log("Image URL for " + code + " : " + image);
}



function pingFromUi() {
  return {
    status: "OK",
    time: new Date().toString()
  };
}

/**
 * Enforce that value becomes a REAL Date object
 * Accepts:
 *  - Date object
 *  - ISO string (YYYY-MM-DD or YYYY-MM-DDTHH:mm:ssZ)
 *  - dd/MM/yyyy (UI fallback)
 * Rejects everything else
 */
function enforceDate_(value, fieldName) {
  if (!value) {
    throw new Error(`❌ ${fieldName} is required`);
  }

  // Already Date
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  // ISO format
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}/.test(value)) {
    const d = new Date(value);
    if (!isNaN(d.getTime())) return d;
  }

  // dd/MM/yyyy fallback
  if (typeof value === "string" && /^\d{2}\/\d{2}\/\d{4}$/.test(value)) {
    const [dd, mm, yyyy] = value.split("/").map(Number);
    const d = new Date(yyyy, mm - 1, dd);
    if (!isNaN(d.getTime())) return d;
  }

  throw new Error(
    `❌ Invalid date for ${fieldName}. Use Date or ISO (YYYY-MM-DD)`
  );
}

function updateSalesOrderWithFiles(payload, files, opts) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    /* ===============================
       VALIDATIONS
    =============================== */
    if (!payload || !payload.products || !payload.products.length) {
      throw new Error("No products in payload.");
    }

    const orderDate = enforceDate_(payload.orderDate, "Order Date");
    const poDate    = payload.poDate ? enforceDate_(payload.poDate, "PO Date") : "";

    /* ===============================
       SHEET
    =============================== */
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Sales Orders");
    if (!sh) throw new Error("Sales Orders sheet not found");

    /* ===============================
       OLD + NEW REF CODE
    =============================== */
    const oldRefCode = String(payload.refCode).trim();
    const newRefCode = getNextRefCodeFromSheet_(sh);

    /* ===============================
       FILE UPLOADS
    =============================== */
    let customerDocURL  = payload.existingCustomerDocUrl || "";
    let agreementDocURL = payload.existingAgreementCopyUrl || "";

    if (Array.isArray(files)) {
      files.forEach(f => {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(f.data),
          f.mimeType,
          f.fileName
        );
        const file = DriveApp.createFile(blob);
        const url = file.getUrl();

        if (f.fieldName === "customerDocs")  customerDocURL  = url;
        if (f.fieldName === "agreementCopy") agreementDocURL = url;
      });
    }

    /* ===============================
       PREPARE PRODUCT IMAGES
    =============================== */
    payload.products = payload.products.map(p => ({
      ...p,
      imageDataUri: fetchImageAsDataUri(p.productImage || p.image || "")
    }));

    /* ===============================
       BUILD ROWS
    =============================== */
    const timestamp = new Date(); // REAL DATE (not string)

    const rows = payload.products.map(p => [
      // ----- MAIN DATA -----
      timestamp,
      newRefCode,
      orderDate,               // ✅ REAL DATE
      payload.fy,
      payload.gstType,
      payload.mobile,
      payload.altMobile,
      payload.email,
      payload.customerName,
      payload.companyName,
      payload.address,
      payload.gstNumber,
      payload.heads,
      payload.user,
      payload.deliverySameAsOrder ? "Yes" : "No",
      payload.delMobile,
      payload.delGstNumber,
      payload.delCompanyName,
      payload.delContactPerson,
      payload.delEmail,
      payload.delAddress,

      // ----- PRODUCT -----
      p.productCode,
      p.productName,
      p.description,
      p.size,
      p.colour,
      p.hsn,
      p.totalQty,
      p.unitValue,
      p.stock,
      p.subTotal,
      p.gstPercent,
      p.gstAmount,
      p.totalPrice,

      // ----- TOTALS -----
      payload.grandTotal,
      payload.poNumber,
      poDate,                  // ✅ REAL DATE
      payload.advanceAmount,
      payload.clientType,
      payload.paymentTerms,
      payload.remarks,
      payload.totalWithoutGst,
      payload.totalCgst,
      payload.totalSgst,
      payload.totalIgst,
      payload.footerGrandTotal,

      // ----- FILES -----
      customerDocURL,
      agreementDocURL,

      // ----- INVOICE + STATUS + REVISION -----
      "",
      "Revised",
      oldRefCode
    ]);

    /* ===============================
       WRITE DATA
    =============================== */
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    /* ===============================
       FORMAT DATE COLUMNS (DISPLAY ONLY)
    =============================== */
    sh.getRange(startRow, 1, rows.length, 1)
      .setNumberFormat("dd/MM/yyyy HH:mm:ss");

    sh.getRange(startRow, 3, rows.length, 1)
      .setNumberFormat("dd/MM/yyyy");

    sh.getRange(startRow, 36, rows.length, 1)
      .setNumberFormat("dd/MM/yyyy");

    /* ===============================
       GENERATE INVOICE
    =============================== */
    const invoiceInfo = generateInvoicePdf_(payload, newRefCode);

    /* ===============================
       WRITE INVOICE URL
    =============================== */
    const invoiceColIndex = rows[0].length - 2;
    sh.getRange(startRow, invoiceColIndex, rows.length, 1)
      .setValues(rows.map(() => [invoiceInfo.pdfUrl || ""]));

    return {
      success: true,
      message: "Sales Order revised successfully.",
      referenceCode: newRefCode,
      oldRefCode: oldRefCode,
      invoiceUrl: invoiceInfo.pdfUrl
    };

  } catch (err) {
    return {
      success: false,
      message: err.message || String(err)
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

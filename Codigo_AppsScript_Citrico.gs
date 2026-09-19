var SHEET_ID = "1FhZb0ssIN0IHl2ROqxNS0ZlwLhKi8HL5oRUoExDxB0M";
var SHEET_NAME = "Ingresos y Gastos";

var HEADERS = [
  "ID","Timestamp","Fecha","Año","Mes","Día","Mes N°",
  "Sucursal","Fuente","Tipo de movimiento",
  "Código SAT","Cuenta SAT","Grupo","Estado financiero",
  "Producto","Proveedor",
  "Efectivo","T. Déb/Cré","TPV","AMEX","Transferencia","Total","Flujo",
  "Folio factura","Notas","Real BanBajío","Pagado","Observaciones",
  "Foto comprobante","Responsable de captura"
];

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var recs = data.records || [];
    
    if(recs.length === 0) return json_({ok:false,error:"Sin registros"});
    
    var sheet = getSheet_();
    var rows = [];
    
    for(var i = 0; i < recs.length; i++) {
      var r = recs[i];
      var id = Utilities.getUuid().slice(0,8);
      rows.push([
        id, new Date(),
        r.fecha||"", r.anio||"", r.mes||"", r.dia||"", r.mes_n||"",
        r.sucursal||"", r.fuente||"", r.tipo_movimiento||"",
        r.codigo_sat||"", r.cuenta_sat||"", r.grupo||"", r.estado_financiero||"",
        r.producto||"", r.proveedor||"",
        num_(r.efectivo), num_(r.debito_credito), num_(r.tpv), num_(r.amex), num_(r.transferencia),
        num_(r.total), num_(r.flujo),
        r.folio_factura||"", r.notas||"", num_(r.real_banbajio), r.pagado||"", r.observaciones||"",
        "", r.capturado_por||""
      ]);
    }
    
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
    
    return json_({ok:true,added:rows.length});
  } catch(err) {
    return json_({ok:false,error:String(err)});
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if(sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1,1,1,HEADERS.length).setFontWeight("bold").setBackground("#BC5A38").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function num_(v) {
  var n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
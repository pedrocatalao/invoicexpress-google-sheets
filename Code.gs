/**
 * This script will connect to invoiceXpress and fetch invoices into a Google sheets document.
 */

/**
 * Access credentials
 */
var API_KEY = "your-api-key";
var SUBDOMAIN = "your-subdomain"

/**
 * Data containers
 */
var GOOD_INVOICES = [];
var FIXED_INVOICES = [];

/**
 * Colors
 */
var RED = '#FF0000';
var GREEN = '#00FF00';
var YELLOW = '#FFFF00';
var GREY = '#666666';
var BRIGHT = '#BBBBBB'
var DARK = '#222222';
var BLACK = '#000000';

/**
 * Creates custom menu entry
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('InvoiceXpress')
      .addItem('Refresh data', 'updateDocuments')
      .addToUi();
}

/**
 * Put request url together
 */
function getEndpointUrl(method, page) {
  var domain = 'app.invoicexpress.com';

  return 'https://' + SUBDOMAIN + '.' + domain + '/' + method + '?api_key=' + API_KEY + '&page=' + page;
}

/**
 * Fetches all documents and stores them in the data container vars
 */
function fetchDocuments() {
  var page = 1;
  var lastPage = 1;
  
  while(page <= lastPage) {
    var url = getEndpointUrl('invoices.json', page++); //set page
    var response = UrlFetchApp.fetch(url); //execute query
    var json = JSON.parse(response);

    lastPage = json.pagination.total_pages;

    for(var i=0; i < json.invoices.length; i++) {
      invoice = json.invoices[i];

      if(invoice.type == 'Invoice' && invoice.status != 'canceled')
        GOOD_INVOICES.push(invoice);

      else
        FIXED_INVOICES.push(invoice);
    }
  }
}

/**
 * Sort documents by id
 */
function sortDocuments() {
  GOOD_INVOICES.sort((a, b) => (a.id > b.id) ? -1 : 1);
  FIXED_INVOICES.sort((a, b) => (a.id > b.id) ? -1 : 1);
}

/**
 * Re-writes the invoice and fixed documents sheets
 */
function updateDocuments() {
  fetchDocuments();
  sortDocuments();
  writeSheet(GOOD_INVOICES, "Invoices");
  writeSheet(FIXED_INVOICES, "Credit Notes and canceled invoices");
}

/**
 * Find existing or create new sheet for invoices
 */
function prepareSheet(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet == null)
    sheet = spreadsheet.insertSheet(sheetName);

  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(),  sheet.getMaxColumns()).setBackground(BLACK).setHorizontalAlignment('left');

  return sheet;
}

/**
 * Create and format header row
 */
function createAndFormatHeader(sheet) {
  var header = ["Number", "Date", "Due on", "Type", "Status", "Customer", "Subtotal", "VAT", "Total", "Download PDF"];
  var headersRange = sheet.getRange(1, 1, 1, header.length);

  headersRange.setValues([header]);
  headersRange.setBackground(DARK);
  headersRange.setFontColor(BRIGHT);
  headersRange.setFontSize(11);
  headersRange.setFontFamily("Oswald");

  sheet.setFrozenRows(1);
}

/**
 * writes a table with a invoice data on each row
 */
function writeSheet(documents, sheetName) {
  var j;
  var sheet = prepareSheet(sheetName);
  createAndFormatHeader(sheet);

  for(var i=0; i < documents.length; i++) {
    var statusColor = getStatusColor(documents[i].status, documents[i].due_date);
    sheet.getRange(i+2,j=1).setValue(documents[i].sequence_number).setFontColor(GREY);
    sheet.getRange(i+2,++j).setValue(documents[i].date).setFontColor(GREY);
    sheet.getRange(i+2,++j).setValue(documents[i].due_date).setFontColor(GREY);
    sheet.getRange(i+2,++j).setValue(documents[i].type).setFontColor(GREY);
    sheet.getRange(i+2,++j).setValue(documents[i].status).setFontColor(statusColor);
    sheet.getRange(i+2,++j).setValue(documents[i].client.name).setFontColor(GREY);
    sheet.getRange(i+2,++j).setValue(documents[i].before_taxes).setFontColor(GREY);
    sheet.getRange(i+2,++j).setValue(documents[i].taxes).setFontColor(GREY);
    sheet.getRange(i+2,++j).setValue(documents[i].total).setFontColor(GREY);
    sheet.getRange(i+2,++j).setFormula('=hyperlink("'+documents[i].permalink+'";"Download '+documents[i].sequence_number+'")');
  }

  autoResizeColumns(sheet, j);
}

/**
 * Resize columns to fit content
 */
function autoResizeColumns(sheet, numberOfColumns) {
  for (var i=1; i < numberOfColumns; i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * get status color
 */
function getStatusColor(status, dueDate) {
  var dateParts = dueDate.split('/');
  var fromDate = new Date(dateParts[2], dateParts[1]-1, dateParts[0]);

  if(status == 'final' && fromDate < new Date()) return RED;
  if(status == 'final') return YELLOW;
  if(status == 'settled') return GREEN;
  if(status == 'canceled') return BRIGHT;
}

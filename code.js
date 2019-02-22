var ss = SpreadsheetApp.getActiveSpreadsheet();
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Menu')
        .addItem('Add To Inventory', 'addToInventory')
        .addItem('Add New Catalog', 'addToCatalog')
        .addSeparator()
        .addToUi();
}
function addToInventory() {
    prepareHTML('addInventory', 'Add To Inventory')
}
function addToCatalog() {
    prepareHTML('addCatalog', 'Create New Catalog')
}
function prepareHTML(file, title) {
    var html = HtmlService.createHtmlOutputFromFile(file)
        .setWidth(600)
        .setHeight(425)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, title);
}
function addCatalog(formObject) {
    var sheet = '';
    var catalogType = formObject.type;
    switch (catalogType) {
        case ('Audio CD'):
            sheet = 'Audio Catalogs'; break;
        case ('Video CD'):
            sheet = 'Video Catalogs'; break;
        default:
            sheet = 'Books Catalogs'
    }
    var catalogs = ss.getSheetByName(sheet);
    var allCatalogs = catalogs.getDataRange().getValues();
    var nextRow = allCatalogs.length;
    var catalogValues = [[new Date(), formObject.title, formObject.speaker, formObject.date, formObject.tag, formObject.location]];
    var nextEntry = catalogs.getRange(++nextRow, 1, 1, catalogValues[0].length);
    nextEntry.setValues(catalogValues);
}
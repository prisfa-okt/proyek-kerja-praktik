function testBacaData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName('Data_Usaha')
  const data = sheet.getDataRange().getValues()
  Logger.log(data)
}

function doGet(e) {
  const halaman = e && e.parameter && e.parameter.page ? e.parameter.page : 'Index'
  
  try {
    const template = HtmlService.createTemplateFromFile(halaman)
    
    // Inject data user ke template kalau ada
    template.nama     = e.parameter.nama     || ''
    template.jabatan  = e.parameter.jabatan  || ''
    template.kota     = e.parameter.kota     || ''
    template.provinsi = e.parameter.provinsi || ''
    template.email    = e.parameter.email    || ''
    template.telepon  = e.parameter.telepon  || ''

    return template.evaluate()
      .setTitle('GIS Usaha Sidoarjo')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  } catch(err) {
    return HtmlService.createHtmlOutput('Error: ' + err)
  }
}

function cekLogin(email, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users')
  const data = sheet.getDataRange().getValues()

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === email && data[i][1].toString() === password) {
      return {
        status  : 'ok',
        nama    : data[i][2],
        jabatan : data[i][3],
        kota    : data[i][4],
        provinsi: data[i][5],
        email   : data[i][0],  // ambil dari kolom pertama
        telepon : data[i][6]
      }
    }
  }
  return { status: 'gagal' }
}

function getPetaUrl() {
  return ScriptApp.getService().getUrl()
}

function getBacaData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Usaha')
  return sheet.getDataRange().getValues()
}

function tambahData(dataInput) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Usaha')
  sheet.appendRow([
    dataInput.idsbr,     // idsbr (kosong, auto)
    dataInput.lat,       // latitude
    dataInput.lon,       // longitude
    dataInput.nama,      // nama_usaha
    dataInput.alamat,    // alamat_usaha
    'JAWA TIMUR',        // nmprov
    'SIDOARJO',          // nmkab
    dataInput.kecamatan, // nmkec
    dataInput.desa,      // nmdesa
    dataInput.kategori,  // kategori_usaha
    'UB'                 // skala_usaha
  ])
}

function editData(noBaris, dataInput) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Usaha')
  sheet.getRange(noBaris, 1).setValue(dataInput.idsbr)     // idsbr
  sheet.getRange(noBaris, 2).setValue(dataInput.lat)       // latitude
  sheet.getRange(noBaris, 3).setValue(dataInput.lon)       // longitude
  sheet.getRange(noBaris, 4).setValue(dataInput.nama)      // nama_usaha
  sheet.getRange(noBaris, 5).setValue(dataInput.alamat)    // alamat_usaha
  sheet.getRange(noBaris, 8).setValue(dataInput.kecamatan) // nmkec
  sheet.getRange(noBaris, 9).setValue(dataInput.desa)      // nmdesa
  sheet.getRange(noBaris, 10).setValue(dataInput.kategori) // kategori_usaha
}

function hapusData(noBaris) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Usaha')
  sheet.deleteRow(noBaris)
}

function getIndexUrl() {
  return ScriptApp.getService().getUrl()
}

// ← TAMBAHKAN DI SINI
function getPetaHtml(role, nama, jabatan, kota, provinsi, email, telepon) {
  const template = HtmlService.createTemplateFromFile('Peta')
  template.role     = role
  template.nama     = nama
  template.jabatan  = jabatan
  template.kota     = kota
  template.provinsi = provinsi
  template.email    = email
  template.telepon  = telepon
  return template.evaluate().getContent()
}
const TEMPLATE_FILE_ID = "1Z_ayoIjk9EzOVFYqkTxiYgNUQSzT5ZkB88NV3nfJOCQ";
const DESTINATION_FOLDER_ID = "1O2r4M4uLAoDJ-KicjfyNwblBnhgxD24u";
const CURRENCY_SIGN = "Rp";

function toCurrency(num) {
  // Pastikan angka diformat dengan 2 desimal
  var fmt = Number(num).toFixed(2);

  // Pisahkan bagian desimal dan integer
  var parts = fmt.split(".");

  // Tambahkan thousand separator ke bagian integer
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");

  // Gabungkan kembali bagian integer dan desimal dengan simbol mata uang
  return `${CURRENCY_SIGN} ${parts.join(",")}`;
}

// Format datetimes to: YYYY-MM-DD
function toDateFmt(dt_string) {
  var millis = Date.parse(dt_string);
  var date = new Date(millis);
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);

  // Return the date in YYYY-mm-dd format
  return `${year}-${month}-${day}`;
}

// Parse and extract the data submitted through the form.
function parseFormData(values, header) {
  // Set temporary variables to hold prices and data.
  var subtotal = 0;
  var discount = 0;
  var response_data = {};

  // Iterate through all of our response data and add the keys (headers)
  // and values (data) to the response dictionary object.
  for (var i = 0; i < values.length; i++) {
    // Extract the key and value
    var key = header[i];
    var value = values[i];

    // If we have a price, add it to the running subtotal and format it to the
    // desired currency.
    if (key.toLowerCase().includes("price")) {
      subtotal += value;
      value = toCurrency(value);

      // If there is a discount, track it so we can adjust the total later and
      // format it to the desired currency.
    } else if (key.toLowerCase().includes("discount")) {
      discount += value;
      value = toCurrency(value);

      // Format dates
    } else if (key.toLowerCase().includes("date")) {
      value = toDateFmt(value);
    }

    // Add the key/value data pair to the response dictionary.
    response_data[key] = value;
  }

  // Once all data is added, we'll adjust the subtotal and total
  response_data["sub_total"] = toCurrency(subtotal);
  response_data["total"] = toCurrency(subtotal - discount);

  return response_data;
}

// Helper function to inject data into the template
function populateTemplate(document, response_data) {
  // Get the document header and body (which contains the text we'll be replacing).
  var document_header = document.getHeader();
  var document_body = document.getBody();

  // Replace variables in the header
  for (var key in response_data) {
    var match_text = `{{${key}}}`;
    var value = response_data[key];

    // Replace our template with the final values
    document_header.replaceText(match_text, value);
    document_body.replaceText(match_text, value);
  }
}

// Function to populate the template form
function createDocFromForm() {
  // Get active sheet and response data.
  var sheet = SpreadsheetApp.getActiveSheet();
  var last_row = sheet.getLastRow();
  var range = sheet.getDataRange();
  var headers = range.getValues()[0]; // Ambil header dari baris pertama

  // Validasi untuk memastikan header "sub_total" dan "total" hanya ditambahkan sekali
  if (!headers.includes("Sub_Total")) {
    sheet.getRange(1, headers.length + 1).setValue("Sub_Total");
    headers.push("Sub_Total"); // Tambahkan header ke array untuk validasi berikutnya
  }
  if (!headers.includes("Total")) {
    sheet.getRange(1, headers.length + 2).setValue("Total");
    headers.push("Total"); // Tambahkan header ke array untuk validasi berikutnya
  }

  // Ambil data terbaru (baris terakhir) untuk diproses
  var data = sheet.getRange(last_row, 1, 1, headers.length).getValues()[0]; // Baris terakhir

  // Parse form data
  var response_data = parseFormData(data, headers);

  // Ambil nilai subtotal dan total dari hasil parse
  var subtotal = response_data["sub_total"]; // Sudah dalam format Rupiah
  var total = response_data["total"]; // Sudah dalam format Rupiah

  // Tulis subtotal dan total ke kolom yang sesuai pada baris terakhir
  sheet.getRange(last_row, headers.indexOf("Sub_Total") + 1).setValue(subtotal);
  sheet.getRange(last_row, headers.indexOf("Total") + 1).setValue(total);

  // Proses pembuatan dokumen
  var template_file = DriveApp.getFileById(TEMPLATE_FILE_ID);
  var target_folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  var filename = `${response_data["Company Name"]}_${response_data["Invoice Date"]}_${response_data["Invoice Number"]}`;
  var document_copy = template_file.makeCopy(filename, target_folder);

  // Buka dokumen salinan dan isi template
  var document = DocumentApp.openById(document_copy.getId());
  populateTemplate(document, response_data);
  document.saveAndClose();
}

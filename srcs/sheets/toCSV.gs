function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛒 Shopify 导出')
    .addItem('导出为 Shopify CSV', 'downloadShopifyCSV')
    .addToUi();
}

function downloadShopifyCSV() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Test");
    const outputFileName = "shopify_products.csv";
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  // Mapping from your sheet columns to Shopify
  const map = {
    "产品编号": "Handle",
    "中文名称": "Title",
    "产品介绍": "Body (HTML)",
    "重量": "Variant Grams",
    "零售价": "Variant Price",
    "数量": "Variant Inventory Qty",
    "图片": "Image Src"
  };

  const shopifyHeaders = [
    "Handle",
    "Title",
    "Body (HTML)",
    "Variant Grams",
    "Variant Price",
    "Variant Inventory Qty",
    "Image Src",
    "Variant Weight Unit",
    "Variant Inventory Tracker",
    "Variant SKU"
  ];

  const csvRows = [shopifyHeaders.join(",")];

  for (const row of rows) {
    const shopifyRow = {};

    for (const [cn, en] of Object.entries(map)) {
      const idx = headers.indexOf(cn);
      if (idx === -1) continue;

      let value = row[idx];

      if (cn === "产品介绍") {
        value = value
          .replace(/&/g, "&amp;")        // escape special chars
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/\r?\n/g, "<br>"); // convert newlines to br>
        value = `<p>${value}</p>`;       // wrap in a div for structure
      }

      // Escape commas
      if (typeof value === "string" && /[",]/.test(value)) {
         value = '"' + value + '"';
      }

      if (cn === "产品介绍") {
        value = `"${value}"`;       // wrap html in a quota
      }

      shopifyRow[en] = value || "";
    }

    // Add fixed Shopify fields
    shopifyRow["Variant Weight Unit"] = "g";
    shopifyRow["Variant Inventory Tracker"] = "shopify";
    shopifyRow["Variant SKU"] = shopifyRow["Handle"];

    const csvLine = shopifyHeaders.map(h => shopifyRow[h] || "").join(",");
    csvRows.push(csvLine);
  }

  const csvContent = csvRows.join("\n");

  // === Download dialog ===
  const html = HtmlService.createHtmlOutput(`
    <html>
      <body>
        <a id="download" href="data:text/csv;charset=utf-8,${encodeURIComponent(csvContent)}"
           download="${outputFileName}">点击这里下载 Shopify CSV</a>
        <script>
          document.getElementById("download").click();
          google.script.host.close();
        </script>
      </body>
    </html>
  `);
  SpreadsheetApp.getUi().showModalDialog(html, "准备下载...");
}

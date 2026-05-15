let bolData = null;
let priceData = null;

document.getElementById("bolFile").addEventListener("change", (e) => {
  handleBOLFile(e.target.files[0]);
});

document.getElementById("priceFile").addEventListener("change", (e) => {
  handlePriceFile(e.target.files[0]);
});

function handleBOLFile(file) {
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: "array" });
      const sheet = workbook.Sheets["BOL"];

      if (!sheet) {
        showError('TAB "BOL" nicht gefunden!');
        return;
      }

      // Dienstleister aus I10 lesen
      const supplier = sheet["I10"]?.v || "Unbekannt";

      // Daten ab Zeile 20 lesen (KORRIGIERT!)
      const rows = [];
      let row = 20;
      while (sheet[`D${row}`]) {
        const oz = sheet[`D${row}`]?.v;
        const menge = sheet[`G${row}`]?.v;

        if (oz) {
          rows.push({
            supplier: supplier,
            oz: String(oz).trim(),
            menge: parseFloat(menge) || 0,
            row: row,
          });
        }
        row++;
      }

      bolData = rows;
      document.getElementById("bolStatus").textContent =
        `✓ ${rows.length} Positionen geladen`;
      document.getElementById("bolIcon").className = "status-icon ok";
      document.getElementById("bolIcon").textContent = "✓";
      document.getElementById("bolText").textContent =
        `${rows.length} Positionen`;

      processData();
    } catch (err) {
      showError("Fehler beim Lesen der BOL-Datei: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function handlePriceFile(file) {
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: "array" });
      const sheet = workbook.Sheets["SAPUI5-Export"];

      if (!sheet) {
        showError('TAB "SAPUI5-Export" nicht gefunden!');
        return;
      }

      // Daten ab Zeile 2 einlesen
      const rows = [];
      let row = 2;
      while (sheet[`F${row}`]) {
        const supplier = sheet[`F${row}`]?.v;
        const shorttext = sheet[`T${row}`]?.v || "";
        const price = sheet[`AA${row}`]?.v;

        // OZ aus Shorttext extrahieren (erste Zahl am Anfang)
        const oz = extractOZ(shorttext);

        if (supplier && oz) {
          rows.push({
            supplier: String(supplier).trim(),
            oz: oz,
            shorttext: extractShorttextDescription(shorttext),
            price: parseFloat(price) || 0,
          });
        }
        row++;
      }

      priceData = rows;
      document.getElementById("priceStatus").textContent =
        `✓ ${rows.length} Artikel geladen`;
      document.getElementById("priceIcon").className = "status-icon ok";
      document.getElementById("priceIcon").textContent = "✓";
      document.getElementById("priceText").textContent =
        `${rows.length} Artikel`;

      processData();
    } catch (err) {
      showError("Fehler beim Lesen der Preisliste: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// OZ aus Shorttext extrahieren (z.B. "11001 Einarbeitung..." → "11001")
function extractOZ(text) {
  if (!text) return "";
  const match = String(text).match(/^(\d+)/);
  return match ? match[1] : "";
}

// Nur die Beschreibung aus Shorttext (z.B. "11001 Einarbeitung..." → "Einarbeitung...")
function extractShorttextDescription(text) {
  if (!text) return "";
  return String(text)
    .replace(/^\d+\s*[-\s]*/, "")
    .trim();
}

function processData() {
  if (!bolData || !priceData) return;

  document.getElementById("loading").style.display = "block";

  setTimeout(() => {
    const results = [];
    let matched = 0;
    let errors = 0;
    let totalSum = 0;

    // Bestellungen mit Preisen matchen
    bolData.forEach((order) => {
      if (order.menge <= 0) return;

      const priceItem = priceData.find(
        (p) => p.supplier === order.supplier && p.oz === order.oz,
      );

      if (priceItem) {
        const gesamtpreis = order.menge * priceItem.price;
        totalSum += gesamtpreis;
        matched++;

        results.push({
          supplier: order.supplier,
          oz: order.oz,
          shorttext: priceItem.shorttext,
          menge: order.menge,
          price: priceItem.price,
          total: gesamtpreis,
          error: false,
        });
      } else {
        errors++;
        results.push({
          supplier: order.supplier,
          oz: order.oz,
          shorttext: "⚠️ Nicht gefunden",
          menge: order.menge,
          price: 0,
          total: 0,
          error: true,
        });
      }
    });

    // Ergebnisse anzeigen
    displayResults(results, matched, errors, totalSum);

    document.getElementById("loading").style.display = "none";
    document.getElementById("resultsContainer").style.display = "block";
  }, 500);
}

function displayResults(results, matched, errors, totalSum) {
  const tbody = document.getElementById("resultsBody");
  tbody.innerHTML = "";

  results.forEach((item) => {
    const row = document.createElement("tr");
    if (item.error) row.className = "error-row";

    row.innerHTML = `
      <td>${item.supplier}</td>
      <td>${item.oz}</td>
      <td>${item.shorttext}</td>
      <td class="number">${item.menge.toFixed(2)}</td>
      <td class="number">${item.price.toFixed(2).replace(".", ",")} €</td>
      <td class="number"><strong>${item.total.toFixed(2).replace(".", ",")} €</strong></td>
    `;

    tbody.appendChild(row);
  });

  // Summary aktualisieren
  document.getElementById("countOrders").textContent = results.length;
  document.getElementById("countMatched").textContent = matched;
  document.getElementById("countErrors").textContent = errors;
  document.getElementById("totalSum").textContent =
    totalSum.toFixed(2).replace(".", ",") + " €";
}

// PDF Export
function exportPDF() {
  const html = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h1 style="text-align: center; color: #333; margin-bottom: 30px;">
        📊 Bestellübersicht
      </h1>
      
      <div style="margin-bottom: 20px; color: #666;">
        <p><strong>Erstellt am:</strong> ${new Date().toLocaleDateString("de-DE")}</p>
      </div>
      
      <table style="width: 100%; border-collapse: collapse; margin-bottom: 30px;">
        <thead>
          <tr style="background: #f8f9fa;">
            <th style="border: 1px solid #dee2e6; padding: 10px; text-align: left;">Item-ID</th>
            <th style="border: 1px solid #dee2e6; padding: 10px; text-align: left;">Kurztext</th>
            <th style="border: 1px solid #dee2e6; padding: 10px; text-align: right;">Menge</th>
            <th style="border: 1px solid #dee2e6; padding: 10px; text-align: right;">Einzelpreis</th>
            <th style="border: 1px solid #dee2e6; padding: 10px; text-align: right;">Gesamtpreis</th>
          </tr>
        </thead>
        <tbody>
          ${Array.from(document.querySelectorAll("#resultsBody tr"))
            .map((row) => {
              const cells = row.querySelectorAll("td");
              const isError = row.classList.contains("error-row");
              const bgColor = isError ? "#ffebee" : "white";
              return `
                <tr style="background: ${bgColor};">
                  <td style="border: 1px solid #dee2e6; padding: 10px;">${cells[1].textContent}</td>
                  <td style="border: 1px solid #dee2e6; padding: 10px;">${cells[2].textContent}</td>
                  <td style="border: 1px solid #dee2e6; padding: 10px; text-align: right;">${cells[3].textContent}</td>
                  <td style="border: 1px solid #dee2e6; padding: 10px; text-align: right;">${cells[4].textContent}</td>
                  <td style="border: 1px solid #dee2e6; padding: 10px; text-align: right; font-weight: bold;">${cells[5].textContent}</td>
                </tr>
              `;
            })
            .join("")}
        </tbody>
      </table>
      
      <div style="text-align: right; padding: 20px; background: #f0f7ff; border-radius: 4px;">
        <div style="font-size: 18px; font-weight: bold; color: #28a745;">
          Gesamtsumme: ${document.getElementById("totalSum").textContent}
        </div>
      </div>
    </div>
  `;

  const opt = {
    margin: 10,
    filename:
      "Bestellungsübersicht_" +
      new Date().toISOString().split("T")[0] +
      ".pdf",
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { orientation: "portrait", unit: "mm", format: "a4" },
  };

  html2pdf().set(opt).from(html).save();
}

// CSV Export
function exportCSV() {
  const tbody = document.getElementById("resultsBody");
  const rows = tbody.querySelectorAll("tr");

  // CSV Header (ohne Dienstleister) - mit Semikolon
  let csv = "OZ;Shorttext;Menge;Einzelpreis;Gesamtpreis\n";

  rows.forEach((row) => {
    const cells = row.querySelectorAll("td");

    const oz = cells[1].textContent.trim();
    const shorttext = cells[2].textContent.trim();
    const menge = cells[3].textContent.trim();
    const preis = cells[4].textContent.trim();
    const gesamt = cells[5].textContent.trim();

    // Semikolon als Trennzeichen
    csv += `${oz};"${shorttext}";${menge};${preis};${gesamt}\n`;
  });

  // Gesamtsumme hinzufügen
  const totalSum = document.getElementById("totalSum").textContent;
  csv += `\nGESAMTSUMME;;;${totalSum}`;

  // UTF-8 BOM hinzufügen (WICHTIG für Umlaute!)
  const BOM = "\uFEFF";
  const blob = new Blob([BOM + csv], {
    type: "text/csv;charset=utf-8;",
  });

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download =
    "Bestellungen_" + new Date().toISOString().split("T")[0] + ".csv";
  link.click();
}

// In Zwischenablage kopieren
function copyToClipboard() {
  const tbody = document.getElementById("resultsBody");
  const rows = tbody.querySelectorAll("tr");

  let text =
    "Dienstleister\tOZ\tShorttext\tMenge\tEinzelpreis\tGesamtpreis\n";

  rows.forEach((row) => {
    const cells = row.querySelectorAll("td");
    const cols = [
      cells[0].textContent,
      cells[1].textContent,
      cells[2].textContent,
      cells[3].textContent,
      cells[4].textContent,
      cells[5].textContent,
    ];
    text += cols.join("\t") + "\n";
  });

  // Gesamtsumme hinzufügen
  const totalSum = document.getElementById("totalSum").textContent;
  text += "\nGESAMTSUMME:\t" + totalSum;

  // In Zwischenablage kopieren
  navigator.clipboard
    .writeText(text)
    .then(() => {
      alert("✓ Daten in Zwischenablage kopiert!");
    })
    .catch(() => {
      alert("Fehler beim Kopieren");
    });
}

function showError(message) {
  const errorEl = document.getElementById("errorMessage");
  errorEl.textContent = "❌ " + message;
  errorEl.style.display = "block";
  setTimeout(() => {
    errorEl.style.display = "none";
  }, 5000);
}

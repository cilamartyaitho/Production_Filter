(function () {
  let filteredDataForExport = [];
  let FileName = '';

  function initSalableProduction() {
    const DOM = {
      monthSelect: document.getElementById("selectmonth"),
      startDate: document.getElementById("startDatecalender"),
      endDate: document.getElementById("endDatecalender"),
      output: document.getElementById("salableOutput"),
      exportBtn: document.getElementById("salableExportBtn")
    };

    const months = ["January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"];

    const sanitizeKey = (key) => key?.trim()?.toUpperCase().replace(/\s+/g, ' ') || '';

    const parseUserDateInput = (d) => {
      if (!d) return null;
      const [a, b, c] = d.split("-");
      return a.length === 4 ? new Date(`${a}-${b}-${c}`) : new Date(`${c}-${b}-${a}`);
    };

    const stripTime = (date) => date instanceof Date ? new Date(date.getFullYear(), date.getMonth(), date.getDate()) : null;

    const formatDateToDDMMYYYY = (date) => {
      if (!(date instanceof Date)) return '';
      const pad = (n) => String(n).padStart(2, '0');
      return `${pad(date.getDate())}-${pad(date.getMonth() + 1)}-${date.getFullYear()}`;
    };

    async function loadAndFilterData() {
      const month = DOM.monthSelect.value.trim();
      const startStr = DOM.startDate.dataset.displayValue || '';
      const endStr = DOM.endDate.dataset.displayValue || '';

      DOM.output.innerHTML = "Loading...";
      DOM.exportBtn.style.display = "none";
      filteredDataForExport = [];
      FileName = month;

      if (!month || !startStr) {
        alert("Please enter both month and start date.");
        return;
      }

      const filePath = `data/Production_Filled_${month}.xlsx`;

      try {
        const response = await fetch(filePath);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
        const sheet = workbook.Sheets["Total_Salable_Prod"];

        if (!sheet) {
          showError("Sheet named 'Total_Salable_Prod' not found in Excel file.");
          return;
        }

        const rawData = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });

        const data = rawData.map(row => {
        const cleaned = {};
        for (const key in row) {
            const value = row[key];

            if (typeof value === "number") {
            cleaned[sanitizeKey(key)] = parseFloat(value.toFixed(3));
            } else {
            cleaned[sanitizeKey(key)] = value;
            }
        }

        if (cleaned["DATE"] instanceof Date) {
            cleaned.__dateRaw = stripTime(cleaned["DATE"]);
            cleaned["DATE"] = formatDateToDDMMYYYY(cleaned.__dateRaw);
        }

        return cleaned;
        });

        const startDate = parseUserDateInput(startStr);
        const endDate = endStr ? parseUserDateInput(endStr) : startDate;

        if (!startDate || isNaN(startDate)) {
          showError("Invalid start date format.");
          return;
        }

        const start = stripTime(startDate);
        const end = stripTime(endDate);

        const filtered = data.filter(row => {
          const d = row.__dateRaw;
          return d instanceof Date && !isNaN(d) && d >= start && d <= end;
        });

        if (!filtered.length) {
          showMessage("No data found for the selected date(s).");
          return;
        }

        filteredDataForExport = filtered.map(({ __dateRaw, ...rest }) => rest);
        DOM.exportBtn.style.display = "inline-block";
        DOM.output.innerHTML = generateHTMLTable(filtered);

      } catch (err) {
        console.error(err);
        showError(`Could not load file. Make sure "Production_Filled_${month}.xlsx" is in the /data folder and contains a sheet named "Total_Salable_Prod".`);
      }
    }

    function exportToExcel() {
      if (!filteredDataForExport.length) {
        alert("No data to export.");
        return;
      }

      const ws = XLSX.utils.json_to_sheet(filteredDataForExport);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Filtered Salable Data");
      XLSX.writeFile(wb, `Salable_Data_${FileName}.xlsx`);
    }

    function showError(message) {
      DOM.output.innerHTML = `<p style="color:red;">${message}</p>`;
    }

    function showMessage(message) {
      DOM.output.innerHTML = `<p>${message}</p>`;
    }

    function generateHTMLTable(data) {
      const columns = Object.keys(data[0]).filter(key => !key.startsWith("__"));
      let html = "<table border='1' cellpadding='5' cellspacing='0'><tr>";
      columns.forEach(col => html += `<th>${col}</th>`);
      html += "</tr>";

      data.forEach(row => {
        html += "<tr>";
        columns.forEach(col => html += `<td>${row[col] ?? ''}</td>`);
        html += "</tr>";
      });

      html += "</table>";
      return html;
    }

    DOM.exportBtn.addEventListener("click", exportToExcel);

    loadAndFilterData();
  }

  window.initSalableProduction = initSalableProduction;
})();

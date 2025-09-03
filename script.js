let filteredDataForExport = [];
let FileName = '';

function sanitizeKey(key) {
  return key?.trim()?.toUpperCase().replace(/\s+/g, ' ');
}

function parseUserDateInput(d) {
  if (!d) return null;
  const parts = d.split("-");
  if (parts.length !== 3) return null;

  if (parts[0].length === 4) {
    return new Date(`${parts[0]}-${parts[1]}-${parts[2]}`);
  } else {
    return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);

  }
}

function stripTime(date) {
  if (!(date instanceof Date)) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatDateToDDMMYYYY(date) {
  if (!(date instanceof Date)) return '';
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

async function loadAndFilterData() {
  const monthInput = document.getElementById('selectmonth').value.trim();
  const startDateStr = document.getElementById('startDatecalender').dataset.displayValue || '';
  const endDateStr = document.getElementById('endDatecalender').dataset.displayValue || '';
  const output = document.getElementById('output');
  const exportBtn = document.getElementById('exportBtn');
  output.innerHTML = "Loading...";
  exportBtn.style.display = "none"; 
  filteredDataForExport = [];
  FileName = monthInput;
  if (!monthInput || !startDateStr) {
    alert("Please enter both month and start date.");
    return;
  }

  const filePath = `data/Production_Filled_${monthInput}.xlsx`;

  try {
    const response = await fetch(filePath);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
      output.innerHTML = `<p style="color:red;">Sheet named "Sheet1" not found in Excel file.</p>`;
      return;
    }

    const rawData = XLSX.utils.sheet_to_json(sheet, {
      defval: '',
      raw: true
    });

    const data = rawData.map(row => {
      const cleanedRow = {};
      for (let key in row) {
        const cleanKey = sanitizeKey(key);
        cleanedRow[cleanKey] = row[key];
      }

      if (cleanedRow["DATE"] instanceof Date) {
        cleanedRow.__dateRaw = stripTime(cleanedRow["DATE"]);
        cleanedRow["DATE"] = formatDateToDDMMYYYY(cleanedRow.__dateRaw);
      }

      return cleanedRow;
    });

    const startDate = parseUserDateInput(startDateStr);
    const endDate = endDateStr ? parseUserDateInput(endDateStr) : startDate;

    if (!startDate || isNaN(startDate)) {
      output.innerHTML = `<p style="color:red;">Invalid start date format.</p>`;
      return;
    }

    const start = stripTime(startDate);
    const end = stripTime(endDate);

    const filtered = data.filter(row => {
      const d = row.__dateRaw;
      if (!d || isNaN(d)) return false;
      return d >= start && d <= end;
    });

    if (filtered.length === 0) {
      output.innerHTML = `<p>No data found for the selected date(s).</p>`;
      return;
    }

    filteredDataForExport = filtered.map(row => {
      const copy = { ...row };
      delete copy.__dateRaw; 
      return copy;
    });
    exportBtn.style.display = "inline-block";

    const columns = Object.keys(filtered[0]).filter(c => !c.startsWith("__"));
    let html = "<table border='1' cellpadding='5' cellspacing='0'><tr>";
    columns.forEach(col => html += `<th>${col}</th>`);
    html += "</tr>";

    filtered.forEach(row => {
      html += "<tr>";
      columns.forEach(col => {
        html += `<td>${row[col] ?? ''}</td>`;
      });
      html += "</tr>";
    });

    html += "</table>";
    output.innerHTML = html;

  } catch (err) {
    console.error(err);
    output.innerHTML = `<p style="color:red;">Could not load file. Make sure "Production_Filled_${monthInput}.xlsx" is in the /data folder and contains a sheet named "Sheet1".</p>`;
  }
}

function exportToExcel() {
  if (!filteredDataForExport.length) {
    alert("No data to export.");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(filteredDataForExport);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Filtered Data");

  XLSX.writeFile(wb, `Filtered_Data_${FileName}.xlsx`);
}

const months = ["January","February","March",
  "April","May","June","July","August","September",
  "October","November","December"];

const monthselect = document.getElementById("selectmonth");
const startDateEl = document.getElementById("startDatecalender");
const endDateEl = document.getElementById("endDatecalender");

(function populateMonthDropdown() {
  const currentDate = new Date();
  const currentMonth = currentDate.getMonth();
  const currentYear = currentDate.getFullYear();

  months.forEach((month, index) => {
    const option = document.createElement("option");
    option.value = months[index];
    option.textContent = month;
    if (index === currentMonth) option.selected = true;
    monthselect.appendChild(option);
  });

  restrictDatePickers(currentYear, currentMonth);
})();

function restrictDatePickers(year, month) {
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);

  const minDate = firstDay.toISOString().split("T")[0];
  const maxDate = lastDay.toISOString().split("T")[0];

  startDateEl.min = minDate;
  startDateEl.max = maxDate;
  endDateEl.min = minDate;
  endDateEl.max = maxDate;

  if (startDateEl.value < minDate || startDateEl.value > maxDate) {
    startDateEl.value = "";
    startDateEl.dataset.displayValue = "";
  }
  if (endDateEl.value < minDate || endDateEl.value > maxDate) {
    endDateEl.value = "";
    endDateEl.dataset.displayValue = "";
  }
}

function enableDatePicker(inputEl) {
  inputEl.placeholder = "DD-MM-YYYY";

  inputEl.addEventListener("focus", () => {
    inputEl.type = "date"; 
    if (inputEl.dataset.value) {
      inputEl.value = inputEl.dataset.value;
    }

    if (typeof inputEl.showPicker === "function") {
      inputEl.showPicker();
    }
  });

  inputEl.addEventListener("blur", () => {
    if (inputEl.value) {
      const [y, m, d] = inputEl.value.split("-");
      const formatted = `${d}-${m}-${y}`;
      inputEl.dataset.value = inputEl.value;
      inputEl.dataset.displayValue = formatted;
      inputEl.type = "text";
      inputEl.value = formatted;
    } else {
      inputEl.type = "text";
      inputEl.value = "";
    }
  });
}

enableDatePicker(startDateEl);
enableDatePicker(endDateEl);

monthselect.addEventListener("change", () => {
  const year = new Date().getFullYear();
  const selectedMonthIndex = months.indexOf(monthselect.value);
  if (selectedMonthIndex >= 0) {
    restrictDatePickers(year, selectedMonthIndex);
  }
});

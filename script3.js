async function showSpinningProd() {
  const month = document.getElementById("selectmonth").value;
  const startDateInput = document.getElementById("startDatecalender").value; 
  const endDateInput = document.getElementById("endDatecalender").value; 

  if (!month || !startDateInput) {
    alert("Please select month and start date.");
    return;
  }

  const filePath = `data/Production_Filled_${month}.xlsx`;

  try {
    const response = await fetch(filePath);
    if (!response.ok) throw new Error("Excel file not found!");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });

    const sheet = workbook.Sheets["Spinning_Prod"];
    if (!sheet) throw new Error("Spinning_Prod sheet not found!");

    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const formatCell = d => {
      if (d instanceof Date) {
        const day = String(d.getDate()).padStart(2, "0");
        const month = String(d.getMonth() + 1).padStart(2, "0");
        const year = d.getFullYear();
        return `${day}-${month}-${year}`;
      }
      return d ? d.toString().trim() : "";
    };

    const parseDate = str => {
      if (!str) return null;
      const parts = str.split("-");
      if (parts[0].length === 4) {
        return new Date(parts[0], parts[1] - 1, parts[2]);
      } 
      else {
        return new Date(parts[2], parts[1] - 1, parts[0]);
      }
    };

    const headers = jsonData[0].map(formatCell);
    const startDateObj = parseDate(startDateInput);
    const endDateObj = parseDate(endDateInput);

    const startDateIndex = headers.findIndex(h => {
      const hd = parseDate(h);
      return hd && hd.getTime() === startDateObj.getTime();
    });

    if (startDateIndex === -1) {
      alert(`Start date ${startDateInput} not found in the sheet!`);
      return;
    }

    let endDateIndex = -1;

    if (endDateObj) {
      endDateIndex = headers.findIndex(h => {
        const hd = parseDate(h);
        return hd && hd.getTime() === endDateObj.getTime();
      });

      if (endDateIndex === -1) {
        const possibleIndices = headers
          .map((h, i) => ({ date: parseDate(h), index: i }))
          .filter(x => x.date && x.date <= endDateObj && x.index >= startDateIndex);
        if (possibleIndices.length > 0) {
          endDateIndex = possibleIndices[possibleIndices.length - 1].index;
        } else {
          endDateIndex = startDateIndex;
        }
      }
    } else {
      endDateIndex = startDateIndex;
    }


    const filteredHeaders = ["ITEM NAME | DATE"].concat(headers.slice(startDateIndex, endDateIndex + 1));
    const filteredData = jsonData
      .slice(1)
      .map(row => [row[0]].concat(row.slice(startDateIndex, endDateIndex + 1)));

    let html = `<table border="1" cellspacing="0" cellpadding="5"><thead><tr>`;
    filteredHeaders.forEach(cell => {
      html += `<th>${cell}</th>`;
    });
    html += `</tr></thead><tbody>`;

    filteredData.forEach(row => {
      html += `<tr>`;
      row.forEach(cell => {
        html += `<td>${cell !== undefined ? cell : ""}</td>`;
      });
      html += `</tr>`;
    });

    html += `</tbody></table>`;

    document.getElementById("output").innerHTML = html;
    document.getElementById("exportBtn").style.display = "inline-block";
    document.getElementById("defaultOutputView").style.display = "block";
    document.getElementById("salableOutputView").style.display = "none";

  } catch (err) {
    alert("Error: " + err.message);
    console.error(err);
  }
}

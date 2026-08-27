(() => {
  'use strict';

  const state = { file: null, rows: [], headers: [] };
  const fileInput = document.getElementById('fileInput');
  const dropZone = document.getElementById('dropZone');
  const browseButton = document.getElementById('browseButton');
  const removeFile = document.getElementById('removeFile');
  const generateButton = document.getElementById('generateButton');
  const fileName = document.getElementById('fileName');
  const fileMeta = document.getElementById('fileMeta');
  const fileCard = document.getElementById('fileCard');
  const previewSection = document.getElementById('previewSection');
  const previewTable = document.getElementById('previewTable');
  const recordCount = document.getElementById('recordCount');
  const reportTitle = document.getElementById('reportTitle');
  const previewTitle = document.getElementById('previewTitle');
  const previewDate = document.getElementById('previewDate');
  const toast = document.getElementById('toast');

  const aliases = {
    district: ['districtname', 'district', 'district_name', 'name', 'distname', 'districtnametext'],
    totalSchools: ['tschool', 'totalschools', 'total_school', 'total schools', 'totalschool', 'noofschools', 'numberofschools', 'schoolstotal', 'totalnoofschools'],
    mealSchools: ['tmschoolsreported', 'mealsschoolsreported', 'schoolreported', 'schoolsreported', 'tmschools', 'mealschools', 'mdmschoolsreported', 'schoolsreportingmeal', 'tmschoolsreporting', 'tmschoolreport'],
    studentCount: ['tstudentsenrolled', 'studentsenrolled', 'totalstudents', 'students', 'totalstudentsenrolled', 'noofstudents', 'studentsstrength'],
    dailySchools: ['tdschoolsreported', 'dailyschoolsreported', 'dailyschools', 'tdschools', 'tdschoolsreporting', 'tdschoolreport', 'dailyschoolsreporting'],
    mealsServed: ['noofmdmserved', 'noofmealserved', 'mealsserved', 'mealserved', 'noofmdmserved', 'mdmserved', 'mealscount', 'noofmeals'],
    date: ['cdate', 'date', 'reportdate', 'reportingdate', 'reporting_date']
  };

  function normalize(value) { return String(value ?? '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
  function findColumn(headers, names) {
    const normalized = names.map(normalize);
    let found = headers.find((header) => normalized.includes(normalize(header)));
    if (found) return found;
    found = headers.find((header) => { const h = normalize(header); return normalized.some((n) => h.length > 3 && n.length > 3 && (h.includes(n) || n.includes(h))); });
    return found;
  }
  function valueFor(row, headers, names) { const key = findColumn(headers, names); return key ? row[key] : ''; }
  function displayNumber(value) { const number = Number(String(value).replace(/,/g, '')); return Number.isFinite(number) && String(value).trim() !== '' ? String(number) : String(value ?? ''); }
  function numericValue(value) { const number = Number(String(value).replace(/,/g, '')); return Number.isFinite(number) ? number : 0; }
  function percent(part, whole) { const result = numericValue(whole) ? (numericValue(part) / numericValue(whole)) * 100 : 0; return `${result.toFixed(2)}%`; }
  function showToast(message) { toast.textContent = message; toast.classList.add('show'); window.setTimeout(() => toast.classList.remove('show'), 3200); }
  function formatBytes(bytes) { return `${(bytes / 1024).toFixed(1)} KB`; }

  function setFile(file) {
    if (!file) return;
    const allowed = /\.(xlsx|xls|csv)$/i.test(file.name);
    if (!allowed) { showToast('Please choose an Excel or CSV file.'); return; }
    if (file.size > 10 * 1024 * 1024) { showToast('That file is larger than 10 MB.'); return; }
    state.file = file;
    fileName.textContent = file.name;
    fileMeta.textContent = `${formatBytes(file.size)} · Ready to read`;
    fileCard.classList.remove('empty-card');
    removeFile.hidden = false;
    readWorkbook(file);
  }

  function reset() {
    state.file = null; state.rows = []; state.headers = [];
    fileInput.value = ''; fileName.textContent = 'No file selected'; fileMeta.textContent = 'Upload a workbook to begin';
    fileCard.classList.add('empty-card'); removeFile.hidden = true; generateButton.disabled = true; previewSection.hidden = true;
  }

  function readWorkbook(file) {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = window.XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const records = window.XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
        if (!records.length) throw new Error('No rows found');
        state.rows = records; state.headers = Object.keys(records[0]);
        renderPreview(); generateButton.disabled = false;
      } catch (error) {
        reset(); showToast('This workbook could not be read. Please check the first sheet.');
      }
    };
    reader.onerror = () => { reset(); showToast('The file could not be opened.'); };
    reader.readAsArrayBuffer(file);
  }

  function renderPreview() {
    const headers = state.headers;
    const columns = [
      { label: '#', get: (_, index) => index + 1 },
      { label: 'District', get: (row) => valueFor(row, headers, aliases.district) },
      { label: 'Total schools', get: (row) => displayNumber(valueFor(row, headers, aliases.totalSchools)) },
      { label: 'M-School Reported', get: (row) => displayNumber(valueFor(row, headers, aliases.mealSchools)) },
      { label: 'Reporting %', get: (row) => percent(valueFor(row, headers, aliases.mealSchools), valueFor(row, headers, aliases.totalSchools)), percentage: true },
      { label: 'Students enrolled', get: (row) => displayNumber(valueFor(row, headers, aliases.studentCount)) },
      { label: 'Daily schools Reported', get: (row) => displayNumber(valueFor(row, headers, aliases.dailySchools)) },
      { label: 'T-Meal Served', get: (row) => displayNumber(valueFor(row, headers, aliases.mealsServed)) }
    ];
    const sortedRows = state.rows.map(row => {
      const total = valueFor(row, headers, aliases.totalSchools);
      const daily = valueFor(row, headers, aliases.dailySchools);
      return { row, sortKey: numericValue(total) ? (numericValue(daily) / numericValue(total)) * 100 : 0 };
    }).sort((a, b) => b.sortKey - a.sortKey).map(item => item.row);
    previewTable.innerHTML = `<thead><tr>${columns.map((column) => `<th>${column.label}</th>`).join('')}</tr></thead><tbody>${sortedRows.slice(0, 50).map((row, index) => `<tr>${columns.map((column) => `<td class="${column.percentage ? 'percentage' : ''}">${column.get(row, index)}</td>`).join('')}</tr>`).join('')}</tbody>`;
    recordCount.textContent = `${state.rows.length} record${state.rows.length === 1 ? '' : 's'}`;
    previewTitle.textContent = reportTitle.value || 'District report';
    const date = valueFor(state.rows[0], headers, aliases.date);
    previewDate.textContent = date ? String(date) : new Date().toLocaleDateString('en-GB');
    previewSection.hidden = false;
  }

  function colorForPercent(value) {
    const num = parseFloat(value);
    if (num >= 100) return [200, 230, 201];
    if (num >= 85) return [255, 249, 196];
    if (num >= 50) return [179, 229, 252];
    return [255, 205, 210];
  }
  function textColorForPercent(value) {
    const num = parseFloat(value);
    if (num >= 100) return [27, 94, 32];
    if (num >= 85) return [245, 127, 23];
    if (num >= 50) return [0, 0, 0];
    return [183, 28, 28];
  }

  function drawHeader(doc, pageWidth, title, reportDate, districtCount) {
    doc.setFillColor(244, 241, 235); doc.rect(0, 0, pageWidth, 30, 'F');
    doc.setTextColor(16, 42, 67); doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.text(title.toUpperCase(), pageWidth / 2, 9, { align: 'center' });
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(88, 105, 118); doc.text(`${districtCount} districts`, pageWidth / 2, 15, { align: 'center' });
    doc.setFillColor(16, 42, 67); doc.roundedRect(pageWidth / 2 - 42, 18, 84, 8, 1.5, 1.5, 'F');
    doc.setTextColor(255, 255, 255); doc.setFont('helvetica', 'bold'); doc.setFontSize(8.5); doc.text(`Reporting Date: ${reportDate}`, pageWidth / 2, 23.5, { align: 'center' });
  }

  function drawLegend(doc, startY, pageWidth) {
    doc.setFontSize(7.5); doc.setFont('helvetica', 'bold'); doc.setTextColor(16, 42, 67);
    doc.text('Colour reference:', 10, startY);
    const legendItems = [
      { color: [200, 230, 201], label: '100%' },
      { color: [255, 249, 196], label: '85–99%' },
      { color: [179, 229, 252], label: '50–85%' },
      { color: [255, 205, 210], label: 'Below 50% — Needs attention' }
    ];
    let legendX = 38;
    legendItems.forEach((item) => {
      doc.setFillColor(item.color[0], item.color[1], item.color[2]);
      doc.rect(legendX, startY - 3.2, 4, 4, 'F');
      doc.setTextColor(40, 40, 40); doc.setFont('helvetica', 'normal'); doc.setFontSize(7);
      doc.text(item.label, legendX + 6, startY);
      legendX += doc.getTextWidth(item.label) + 18;
    });
  }

  function buildPdf() {
    if (!state.rows.length || !window.jspdf?.jsPDF) { showToast('The PDF tools are still loading. Try again in a moment.'); return; }
    const landscape = document.getElementById('landscapeToggle').checked;
    const headers = state.headers;
    const reportDate = valueFor(state.rows[0], headers, aliases.date) || new Date().toLocaleDateString('en-GB');
    const title = reportTitle.value.trim() || 'District Wise Meal Served Statistics';

    const sortedRows = state.rows.map(row => {
      const total = valueFor(row, headers, aliases.totalSchools);
      const daily = valueFor(row, headers, aliases.dailySchools);
      return { row, sortKey: numericValue(total) ? (numericValue(daily) / numericValue(total)) * 100 : 0 };
    }).sort((a, b) => b.sortKey - a.sortKey).map(item => item.row);

    const body = sortedRows.map((row, index) => {
      const total = valueFor(row, headers, aliases.totalSchools);
      const meal = valueFor(row, headers, aliases.mealSchools);
      const daily = valueFor(row, headers, aliases.dailySchools);
      return [index + 1, valueFor(row, headers, aliases.district), displayNumber(total), displayNumber(meal), percent(meal, total), displayNumber(daily), percent(daily, total), displayNumber(valueFor(row, headers, aliases.mealsServed))];
    });
    const totalSchools = state.rows.reduce((sum, row) => sum + numericValue(valueFor(row, headers, aliases.totalSchools)), 0);
    const totalMeal = state.rows.reduce((sum, row) => sum + numericValue(valueFor(row, headers, aliases.mealSchools)), 0);
    const totalDaily = state.rows.reduce((sum, row) => sum + numericValue(valueFor(row, headers, aliases.dailySchools)), 0);
    const totalServed = state.rows.reduce((sum, row) => sum + numericValue(valueFor(row, headers, aliases.mealsServed)), 0);
    const totalRowIndex = body.length;
    body.push(['', 'TOTAL', displayNumber(totalSchools), displayNumber(totalMeal), percent(totalMeal, totalSchools), displayNumber(totalDaily), percent(totalDaily, totalSchools), displayNumber(totalServed)]);

    const baseFontSize = landscape ? 11 : 9.5;
    const basePadding = landscape ? 3 : 2.4;
    let fontSize = baseFontSize;
    let cellPadding = basePadding;

    let doc = null;
    for (let attempt = 0; attempt < 20; attempt++) {
      doc = new window.jspdf.jsPDF({ orientation: landscape ? 'landscape' : 'portrait', unit: 'mm', format: 'a4' });
      const pageWidth = doc.internal.pageSize.getWidth();
      drawHeader(doc, pageWidth, title, reportDate, state.rows.length);

      doc.autoTable({
        startY: 33,
        head: [['S/N', 'District', ['Total school'], ['M-School Reported'], ['Reporting %'], ['Daily schools Reported'], ['Reporting %'], ['T-Meal served']]],
        body,
        theme: 'grid',
        margin: { left: 10, right: 10, top: 33, bottom: 18 },
        styles: { font: 'helvetica', fontSize, cellPadding, lineColor: [40, 40, 40], lineWidth: .12, textColor: [20, 20, 20], overflow: 'linebreak', halign: 'center' },
        headStyles: { fillColor: [16, 42, 67], textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center', cellPadding: Math.min(cellPadding, 2) },
        columnStyles: { 0: { halign: 'center', cellWidth: 'auto' }, 1: { halign: 'center', cellWidth: 'auto' }, 2: { halign: 'center', cellWidth: 'auto' }, 3: { halign: 'center', cellWidth: 'auto' }, 4: { halign: 'center', cellWidth: 'auto' }, 5: { halign: 'center', cellWidth: 'auto' }, 6: { halign: 'center', cellWidth: 'auto' }, 7: { halign: 'center', cellWidth: 'auto' } },
        pageBreak: 'avoid',
        rowPageBreak: 'avoid',
        didParseCell: (data) => {
          const isPercentCol = data.column.index === 4 || data.column.index === 6;
          const isTotalRow = data.row.index === totalRowIndex;
          if (data.section === 'body' && isPercentCol) {
            const value = parseFloat(data.cell.raw);
            data.cell.styles.fillColor = colorForPercent(value);
            data.cell.styles.textColor = textColorForPercent(value);
            data.cell.styles.fontStyle = 'bold';
          }
         if (isTotalRow) {
  if (!isPercentCol) {
    data.cell.styles.fillColor = [16, 42, 67];
    data.cell.styles.textColor = [255, 255, 255];
  } else {
    data.cell.styles.textColor = [0, 0, 0];
  }
  data.cell.styles.fontStyle = 'bold';
}

        }
      });

      if (doc.getNumberOfPages() === 1) break;
      fontSize = Math.max(3, fontSize * 0.88);
      cellPadding = Math.max(0.4, cellPadding * 0.88);
    }

    while (doc.getNumberOfPages() > 1) doc.deletePage(doc.getNumberOfPages());

    const legendY = doc.lastAutoTable.finalY + 7;
    const pageWidth = doc.internal.pageSize.getWidth();
    drawLegend(doc, legendY, pageWidth);

    doc.save(`${title.replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').toLowerCase() || 'district-report'}.pdf`);
    showToast('Your PDF has been downloaded.');
  }

  browseButton.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (event) => setFile(event.target.files[0]));
  removeFile.addEventListener('click', reset);
  generateButton.addEventListener('click', buildPdf);
  reportTitle.addEventListener('input', () => { previewTitle.textContent = reportTitle.value || 'District report'; });
  ['dragenter', 'dragover'].forEach((eventName) => dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.add('dragging'); }));
  ['dragleave', 'drop'].forEach((eventName) => dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.remove('dragging'); }));
  dropZone.addEventListener('drop', (event) => setFile(event.dataTransfer.files[0]));
})();

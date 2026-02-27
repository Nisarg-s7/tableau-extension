/* global d3 */
/* global tinycolor */
/* global tableau */

const backgroundColor = tinycolor('white');
const palette = ['#5B6FD8', '#D3D3D3', '#4e79a7', '#f28e2c'];

async function GaugeChart (encodedData, encodingMap, width, height, selectedTupleIds, styles) {
  const margin = { top: 60, right: 80, bottom: 60, left: 80 };
  const innerWidth = width - margin.left - margin.right;
  const innerHeight = height - margin.top - margin.bottom;

  let valueKey = null;
  let targetKey = null;

  // 1. MEASURE DETECTION (Sahi tarike se Sales aur Calculation ko dhundna)
  const numericKeys = [];
  if (encodedData.length > 0) {
    const sample = encodedData[0];
    for (const key in sample) {
      if (key === 'tupleId') continue;
      const arr = sample[key];
      if (Array.isArray(arr) && arr.length > 0 && typeof arr[0].value === 'number') {
        numericKeys.push(key);
      }
    }
  }

  if (numericKeys.length > 0) {
    // Agar kisi field mein 'calc', 'target', 'goal', ya 'budget' likha hai to wo target hai
    targetKey = numericKeys.find(k => 
      k.toLowerCase().includes('calc') || 
      k.toLowerCase().includes('target') || 
      k.toLowerCase().includes('goal') || 
      k.toLowerCase().includes('budget')
    );
    
    // Bacha hua doosra field Sales (Value) hai
    valueKey = numericKeys.find(k => k !== targetKey) || numericKeys[0];
    if (!targetKey && numericKeys.length > 1) targetKey = numericKeys[1];
  }

  if (!valueKey) return { viz: null };

  // 2. AGGREGATE DATA
  let totalValue = 0;
  let totalTarget = 0;
  const allTupleIds = [];

  encodedData.forEach(row => {
    totalValue += parseFloat(row[valueKey]?.[0]?.value || 0);
    if (targetKey) totalTarget += parseFloat(row[targetKey]?.[0]?.value || 0);
    if (row.tupleId) allTupleIds.push(row.tupleId);
  });

  // 3. SCALE CALCULATION
  // Agar target negative hai (aapke screenshot ki tarah), to use scale mein count nahi karenge
  let maxVal = Math.max(totalValue, totalTarget);
  let rawMax = maxVal <= 0 ? 1000 : maxVal * 1.2; 
  const pow10 = Math.pow(10, Math.floor(Math.log10(rawMax)));
  let maxScale = Math.ceil(rawMax / (pow10 / 2)) * (pow10 / 2);

  // 4. SVG SETUP
  const svg = d3.create('svg')
    .attr('class', tableau.ClassNameKey.Worksheet)
    .attr('width', width).attr('height', height)
    .attr('viewBox', [0, 0, width, height])
    .style('background', 'white');

  const cx = width / 2;
  const cy = height * 0.65; 
  const radius = Math.min(innerWidth, innerHeight) * 0.52;
  const chartGroup = svg.append('g').attr('transform', `translate(${cx}, ${cy})`);

  const startAngle = -Math.PI * 0.75; 
  const endAngle = Math.PI * 0.75;   
  const totalRange = endAngle - startAngle;

  const valueFraction = Math.min(Math.max(totalValue / maxScale, 0), 1);
  const currentAngle = startAngle + (valueFraction * totalRange);

  const arcGen = d3.arc().innerRadius(radius * 0.72).outerRadius(radius).cornerRadius(6);

  // 5. DRAW ARCS
  // Background
  chartGroup.append('path').attr('d', arcGen({ startAngle, endAngle })).attr('fill', '#E5E5E5');
  // Blue Sales Arc
  chartGroup.append('path').attr('d', arcGen({ startAngle, endAngle: currentAngle })).attr('fill', palette[0]);

  // 6. DRAW TARGET MARKER (Sabse upar dikhega)
  if (targetKey) {
      // Negative target ko 0 position par dikhayenge
      const tFraction = Math.min(Math.max(totalTarget / maxScale, 0), 1);
      const tAngle = startAngle + (tFraction * totalRange);
      const tx = Math.sin(tAngle);
      const ty = -Math.cos(tAngle);

      const targetG = chartGroup.append('g').attr('class', 'target-marker');

      // Orange Thick Line
      targetG.append('line')
        .attr('x1', radius * 0.68 * tx).attr('y1', radius * 0.68 * ty)
        .attr('x2', radius * 1.18 * tx).attr('y2', radius * 1.18 * ty)
        .attr('stroke', '#f28e2c')
        .attr('stroke-width', 8)
        .attr('stroke-linecap', 'round')
        .style('filter', 'drop-shadow(0px 0px 3px rgba(0,0,0,0.4))');

      // Orange dot at the end
      targetG.append('circle')
        .attr('cx', radius * 1.18 * tx).attr('cy', radius * 1.18 * ty)
        .attr('r', 6).attr('fill', '#f28e2c');
  }

  // 7. TICKS & LABELS
  const numLabels = 5;
  for (let i = 0; i <= numLabels; i++) {
      const f = i / numLabels;
      const angle = startAngle + (f * totalRange);
      const x = Math.sin(angle); const y = -Math.cos(angle);
      chartGroup.append('line')
        .attr('x1', radius * 1.05 * x).attr('y1', radius * 1.05 * y)
        .attr('x2', radius * 1.12 * x).attr('y2', radius * 1.12 * y)
        .attr('stroke', '#444').attr('stroke-width', 2);
      chartGroup.append('text')
        .attr('x', radius * 1.32 * x).attr('y', radius * 1.32 * y)
        .attr('text-anchor', 'middle').attr('dominant-baseline', 'middle')
        .style('font-size', '13px').style('fill', '#555')
        .text(d3.format(".2s")(f * maxScale));
  }

  // 8. NEEDLE
  const needleGroup = chartGroup.append('g').attr('transform', `rotate(${(currentAngle * 180 / Math.PI)})`);
  needleGroup.append('path')
    .attr('d', `M -12 0 L 0 -${radius * 0.88} L 12 0 Z`)
    .attr('fill', '#000')
    .style('filter', 'drop-shadow(2px 2px 4px rgba(0,0,0,0.3))');
  chartGroup.append('circle').attr('r', 18).attr('fill', '#000');
  chartGroup.append('circle').attr('r', 8).attr('fill', '#333');

  // 9. CENTER TEXT
  const textCenterY = radius * 0.25; 
  chartGroup.append('text').attr('y', textCenterY).attr('text-anchor', 'middle')
    .style('font-size', '44px').style('font-weight', 'bold').style('fill', palette[0])
    .text('$' + d3.format(',.0f')(totalValue));

  if (targetKey) {
      chartGroup.append('text').attr('y', textCenterY + 40).attr('text-anchor', 'middle')
        .style('font-size', '18px').style('font-weight', '600').style('fill', '#666')
        .text('Target: ' + d3.format("$,.0f")(totalTarget));
  }

  chartGroup.append('text').attr('y', textCenterY + 65).attr('text-anchor', 'middle')
    .style('font-size', '14px').style('fill', '#999')
    .text(`${(valueFraction * 100).toFixed(1)}% of scale`);

  return { viz: svg.node(), allTupleIds };
}

// --- BOILERPLATE RENDER LOGIC ---
async function renderViz (rawData, encodingMap, selectedMarksIds, styles) {
  const encodedData = getEncodedData(rawData, encodingMap);
  const content = document.getElementById('content');
  content.innerHTML = '';
  const result = await GaugeChart(encodedData, encodingMap, content.offsetWidth, content.offsetHeight, selectedMarksIds, styles);
  if (result.viz) content.appendChild(result.viz);
  return result;
}

window.onload = function() {
  tableau.extensions.initializeAsync().then(async () => {
    const worksheet = getWorksheet();
    let summaryData = {}, encodingMap = {}, selectedMarks = new Map();
    const styles = tableau.extensions.environment.workbookFormatting?.formattingSheets?.find(x => x.classNameKey === 'tableau-worksheet')?.cssProperties;

    const update = async () => {
      [summaryData, encodingMap] = await Promise.all([getSummaryDataTable(worksheet), getEncodingMap()]);
      selectedMarks = await getSelection(worksheet, summaryData);
      await renderViz(summaryData, encodingMap, selectedMarks, styles);
    };

    worksheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, update);
    worksheet.addEventListener(tableau.TableauEventType.FilterChanged, update);

    document.body.addEventListener('click', async (e) => {
        const data = d3.select(e.target).datum();
        if (data && data.allTupleIds) {
            worksheet.selectTuplesAsync(data.allTupleIds, tableau.SelectOptions.Simple);
        } else {
            worksheet.releaseSelectionAsync();
        }
    });

    await update();
  });
};

async function getEncodingMap () {
    const worksheet = getWorksheet();
    const visualSpec = await worksheet.getVisualSpecificationAsync();
    const encodingMap = {};
    if (visualSpec.activeMarksSpecificationIndex < 0) return encodingMap;
    const marksCard = visualSpec.marksSpecifications[visualSpec.activeMarksSpecificationIndex];
    for (const encoding of marksCard.encodings) {
      if (!encodingMap[encoding.id]) encodingMap[encoding.id] = [];
      encodingMap[encoding.id].push(encoding.field);
    }
    return encodingMap;
}
async function getSummaryDataTable(worksheet) {
    let rows = [];
    const reader = await worksheet.getSummaryDataReaderAsync(undefined, { ignoreSelection: true });
    for (let i = 0; i < reader.pageCount; i++) rows = rows.concat(convertToListOfNamedRows(await reader.getPageAsync(i)));
    await reader.releaseAsync();
    return rows;
}
function convertToListOfNamedRows(dataTablePage) {
    const rows = []; const columns = dataTablePage.columns; const data = dataTablePage.data;
    for (let i = 0; i < data.length; ++i) {
      const row = {};
      for (let j = 0; j < columns.length; ++j) row[columns[j].fieldName] = data[i][columns[j].index];
      row.tupleId = i + 1; rows.push(row);
    }
    return rows;
}
function getEncodedData(data, encodingMap) {
    const encodedData = []; let tupleId = 1;
    for (const row of data) {
      const encodedRow = {};
      for (const encName in encodingMap) {
        encodedRow[encName] = [];
        for (const field of encodingMap[encName]) encodedRow[encName].push(row[field.name]);
      }
      encodedRow.tupleId = tupleId++; encodedData.push(encodedRow);
    }
    return encodedData;
}
function getWorksheet() {
    return tableau.extensions.worksheetContent ? tableau.extensions.worksheetContent.worksheet : tableau.extensions.dashboardContent.dashboard.worksheets[0];
}
async function getSelection(worksheet, allMarks) {
    try {
        const selectedMarks = await worksheet.getSelectedMarksAsync();
        if (!selectedMarks.data[0]) return new Map();
        const columns = selectedMarks.data[0].columns;
        const selectedMarksIds = new Map();
        for (const sm of convertToListOfNamedRows(selectedMarks.data[0])) {
            let key = columns.map(col => sm[col.fieldName].value).join('|');
            allMarks.forEach((m, idx) => {
                if (columns.map(col => m[col.fieldName].value).join('|') === key) selectedMarksIds.set(idx + 1);
            });
        }
        return selectedMarksIds;
    } catch (e) { return new Map(); }
}
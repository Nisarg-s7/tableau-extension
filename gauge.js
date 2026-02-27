/* global d3 */
/* global tinycolor */
/* global tableau */

const backgroundColor = tinycolor('white');
const palette = ['#5B6FD8', '#D3D3D3', '#4e79a7', '#f28e2c'];

// 1. MAIN GAUGE CHART FUNCTION
async function GaugeChart (encodedData, encodingMap, width, height, selectedTupleIds, styles) {
  const margin = { top: 60, right: 80, bottom: 60, left: 80 };
  const innerWidth = width - margin.left - margin.right;
  const innerHeight = height - margin.top - margin.bottom;

  let valueKey = null;
  let targetKey = null;

  // Measures Detect karna
  const numericKeys = (data => {
    if (!data || data.length === 0) return [];
    const sample = data[0];
    const keys = [];
    for (const key in sample) {
        if (key === 'tupleId') continue;
        const arr = sample[key];
        if (Array.isArray(arr) && arr.length > 0 && typeof arr[0].value === 'number') keys.push(key);
    }
    return keys;
  })(encodedData);

  if (numericKeys.length > 0) {
      const foundTarget = numericKeys.find(k => k.toLowerCase().includes('target') || k.toLowerCase().includes('calc'));
      if (foundTarget && numericKeys.length > 1) {
          targetKey = foundTarget;
          valueKey = numericKeys.find(k => k !== foundTarget);
      } else {
          valueKey = numericKeys[0];
          if (numericKeys.length > 1) targetKey = numericKeys[1];
      }
  }

  if (!valueKey) return { viz: null };

  // Data Aggregate karna
  let totalValue = 0;
  let totalTarget = 0;
  const allTupleIds = [];
  encodedData.forEach(row => {
    totalValue += parseFloat(row[valueKey]?.[0]?.value || 0);
    if (targetKey) totalTarget += parseFloat(row[targetKey]?.[0]?.value || 0);
    if (row.tupleId) allTupleIds.push(row.tupleId);
  });

  // Dynamic Scale Calculation
  let maxVal = Math.max(totalValue, totalTarget);
  let rawMax = maxVal === 0 ? 1000 : maxVal * 1.2; 
  const pow10 = Math.pow(10, Math.floor(Math.log10(rawMax)));
  let maxScale = Math.ceil(rawMax / (pow10 / 2)) * (pow10 / 2);

  // SVG Setup
  const svg = d3.create('svg')
    .attr('class', tableau.ClassNameKey.Worksheet)
    .attr('width', width)
    .attr('height', height)
    .attr('viewBox', [0, 0, width, height])
    .style('background', 'white');

  const cx = width / 2;
  const cy = height * 0.6; 
  const radius = Math.min(innerWidth, innerHeight) * 0.5;
  const chartGroup = svg.append('g').attr('transform', `translate(${cx}, ${cy})`);

  // Angles logic (270 Degree)
  const startAngle = -Math.PI * 0.75; 
  const endAngle = Math.PI * 0.75;   
  const totalRange = endAngle - startAngle;
  const valueFraction = Math.min(Math.max(totalValue / maxScale, 0), 1);
  const currentAngle = startAngle + (valueFraction * totalRange);

  const arcGen = d3.arc().innerRadius(radius * 0.7).outerRadius(radius).cornerRadius(5);

  // Arcs Draw karna
  chartGroup.append('path').attr('d', arcGen({ startAngle, endAngle })).attr('fill', '#D3D3D3');
  chartGroup.append('path').attr('d', arcGen({ startAngle, endAngle: currentAngle })).attr('fill', palette[0]);

  // Ticks & Labels (Synchronized)
  const numLabels = 5;
  for (let i = 0; i <= numLabels; i++) {
      const f = i / numLabels;
      const angle = startAngle + (f * totalRange);
      const x = Math.sin(angle); const y = -Math.cos(angle);
      chartGroup.append('line')
        .attr('x1', radius * 1.02 * x).attr('y1', radius * 1.02 * y)
        .attr('x2', radius * 1.12 * x).attr('y2', radius * 1.12 * y)
        .attr('stroke', '#333').attr('stroke-width', 2);
      chartGroup.append('text')
        .attr('x', radius * 1.3 * x).attr('y', radius * 1.3 * y)
        .attr('text-anchor', 'middle').attr('dominant-baseline', 'middle')
        .style('font-size', '12px').text(d3.format(".2s")(f * maxScale));
  }

  // Target Marker Line
  if (targetKey && totalTarget > 0) {
      const tAngle = startAngle + (Math.min(totalTarget / maxScale, 1) * totalRange);
      chartGroup.append('line')
        .attr('x1', radius * 0.65 * Math.sin(tAngle)).attr('y1', radius * 0.65 * -Math.cos(tAngle))
        .attr('x2', radius * 1.15 * Math.sin(tAngle)).attr('y2', radius * 1.15 * -Math.cos(tAngle))
        .attr('stroke', '#f28e2c').attr('stroke-width', 5).attr('stroke-linecap', 'round');
  }

  // Needle
  const needleGroup = chartGroup.append('g').attr('transform', `rotate(${(currentAngle * 180 / Math.PI)})`);
  needleGroup.append('path').attr('d', `M -10 0 L 0 -${radius * 0.85} L 10 0 Z`).attr('fill', '#000');
  chartGroup.append('circle').attr('r', 15).attr('fill', '#000');

  // --- CENTER TEXT DISPLAY (Actual, Target, & Percent) ---
  const textCenterY = radius * 0.2; 

  // 1. Sales Value (Main)
  chartGroup.append('text')
    .attr('y', textCenterY)
    .attr('text-anchor', 'middle')
    .style('font-size', '40px')
    .style('font-weight', 'bold')
    .style('fill', palette[0])
    .text('$' + d3.format(',.0f')(totalValue));

  // 2. Target Value (Below Sales)
  if (targetKey && totalTarget > 0) {
      chartGroup.append('text')
        .attr('y', textCenterY + 35)
        .attr('text-anchor', 'middle')
        .style('font-size', '16px')
        .style('font-weight', '500')
        .style('fill', '#666')
        .text('Target: ' + d3.format("$,.0f")(totalTarget));
  }

  // 3. Percentage of Scale
  chartGroup.append('text')
    .attr('y', textCenterY + 60)
    .attr('text-anchor', 'middle')
    .style('font-size', '12px')
    .style('fill', '#999')
    .text(`${(valueFraction * 100).toFixed(1)}% of scale`);

  // Interaction Element
  const interactionElement = chartGroup.append('circle')
    .attr('r', radius).attr('fill', 'transparent').style('cursor', 'pointer');
  interactionElement.datum({ allTupleIds, value: totalValue });

  return { viz: svg.node(), interactionElement, allTupleIds };
}

// 2. RENDER LOGIC
async function renderViz (rawData, encodingMap, selectedMarksIds, styles) {
  const encodedData = getEncodedData(rawData, encodingMap);
  const content = document.getElementById('content');
  content.innerHTML = '';
  const result = await GaugeChart(encodedData, encodingMap, content.offsetWidth, content.offsetHeight, selectedMarksIds, styles);
  if (result.viz) content.appendChild(result.viz);
  return result;
}

// 3. INITIALIZATION
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

// 4. HELPER FUNCTIONS
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
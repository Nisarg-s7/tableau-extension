/* global d3 */
/* global tinycolor */
/* global tableau */

const backgroundColor = tinycolor('white');
const palette = ['#5B8DB8', '#E8E8E8', '#4e79a7', '#e15759'];

async function DonutGaugeChart (encodedData, encodingMap, width, height, selectedTupleIds, styles) {
  const margin = { top: 80, right: 80, bottom: 100, left: 80 };
  const innerWidth = width - margin.left - margin.right;
  const innerHeight = height - margin.top - margin.bottom;
  const radius = Math.min(innerWidth, innerHeight) / 2;

  // 1. Get actual value and target from data
  let actualValue = 0;
  let targetValue = 0;
  let dimension = '';

  encodedData.forEach(row => {
    // Dimension (Region)
    if (row.category && row.category[0]) {
      dimension = row.category[0].formattedValue || '';
    }
    // Value field (SUM(Sales))
    if (row.measure && row.measure[0]) {
      actualValue += parseFloat(row.measure[0].value) || 0;
    }
    // Target field (SUM(Calculation))
    if (row.target && row.target[0]) {
      targetValue += parseFloat(row.target[0].value) || 0;
    }
  });

  // If no target, set a reasonable max
  if (targetValue === 0) {
    targetValue = actualValue * 1.5;
  }

  // Calculate max scale (round up nicely)
  const maxScale = Math.ceil(targetValue / 1000000) * 1000000; // Round to nearest million
  const midScale = maxScale / 2;

  // Calculate percentage based on max scale
  const percentage = Math.min(100, (actualValue / maxScale) * 100);

  console.log('Actual:', actualValue, 'Target:', targetValue, 'Max Scale:', maxScale, 'Percentage:', percentage.toFixed(1) + '%');

  // 2. Create SVG
  const svg = d3.create('svg')
    .attr('class', tableau.ClassNameKey.Worksheet)
    .attr('width', width)
    .attr('height', height)
    .attr('viewBox', [0, 0, width, height])
    .attr('style', 'max-width: 100%; height: auto;')
    .attr('font-family', styles?.fontFamily || 'Segoe UI, Arial')
    .attr('font-size', styles?.fontSize || '12px');

  const chartGroup = svg.append('g')
    .attr('transform', `translate(${width / 2},${height / 2 + 20})`);

  // 3. Create donut arc (semicircle - 180 degrees)
  const outerRadius = radius * 0.9;
  const innerRadius = radius * 0.65;

  // Background (gray) - full semicircle
  const backgroundArc = d3.arc()
    .innerRadius(innerRadius)
    .outerRadius(outerRadius)
    .startAngle(-Math.PI / 2)
    .endAngle(Math.PI / 2);

  chartGroup.append('path')
    .attr('d', backgroundArc)
    .attr('fill', '#E8E8E8')
    .attr('opacity', 0.4);

  // Blue arc (actual value)
  const actualAngle = -Math.PI / 2 + (percentage / 100) * Math.PI;
  
  const actualArc = d3.arc()
    .innerRadius(innerRadius)
    .outerRadius(outerRadius)
    .startAngle(-Math.PI / 2)
    .endAngle(actualAngle);

  chartGroup.append('path')
    .attr('d', actualArc)
    .attr('fill', '#5B8DB8')
    .attr('opacity', 0.9);

  // 4. Draw scale labels (0.0M, middle, max)
  const scalePositions = [
    { value: 0, angle: -Math.PI / 2, label: '0.0M' },
    { value: midScale, angle: 0, label: formatValue(midScale) },
    { value: maxScale, angle: Math.PI / 2, label: formatValue(maxScale) }
  ];

  scalePositions.forEach(pos => {
    const labelRadius = outerRadius * 1.2;
    const x = labelRadius * Math.cos(pos.angle);
    const y = labelRadius * Math.sin(pos.angle);

    chartGroup.append('text')
      .attr('x', x)
      .attr('y', y)
      .attr('text-anchor', 'middle')
      .attr('dy', pos.angle === 0 ? '-0.5em' : '0.5em')
      .style('font-size', '13px')
      .style('font-weight', '500')
      .style('fill', '#666')
      .text(pos.label);
  });

  // 5. Draw tick marks
  for (let i = 0; i <= 100; i += 10) {
    const angle = -Math.PI / 2 + (i / 100) * Math.PI;
    const tickInner = outerRadius * 0.95;
    const tickOuter = i % 20 === 0 ? outerRadius * 1.05 : outerRadius * 1.0;
    
    chartGroup.append('line')
      .attr('x1', tickInner * Math.cos(angle))
      .attr('y1', tickInner * Math.sin(angle))
      .attr('x2', tickOuter * Math.cos(angle))
      .attr('y2', tickOuter * Math.sin(angle))
      .attr('stroke', '#999')
      .attr('stroke-width', i % 20 === 0 ? 2 : 1);
  }

  // 6. Draw needle
  const needleAngle = -Math.PI / 2 + (percentage / 100) * Math.PI;
  const needleLength = outerRadius * 0.85;

  // Needle line
  chartGroup.append('line')
    .attr('x1', 0)
    .attr('y1', 0)
    .attr('x2', needleLength * Math.cos(needleAngle))
    .attr('y2', needleLength * Math.sin(needleAngle))
    .attr('stroke', '#333')
    .attr('stroke-width', 3)
    .attr('stroke-linecap', 'round')
    .style('filter', 'drop-shadow(1px 1px 2px rgba(0,0,0,0.3))');

  // Needle base
  chartGroup.append('circle')
    .attr('r', 10)
    .attr('fill', '#333');

  chartGroup.append('circle')
    .attr('r', 5)
    .attr('fill', '#666');

  // 7. Add "TOTAL" label above value
  chartGroup.append('text')
    .attr('y', 15)
    .attr('text-anchor', 'middle')
    .style('font-size', '12px')
    .style('font-weight', '500')
    .style('fill', '#666')
    .text('TOTAL');

  // 8. Center actual value (below TOTAL label)
  chartGroup.append('text')
    .attr('y', 40)
    .attr('text-anchor', 'middle')
    .style('font-size', '36px')
    .style('font-weight', 'bold')
    .style('fill', '#333')
    .text(d3.format(',.0f')(actualValue));

  // 9. Target value (below actual value)
  chartGroup.append('text')
    .attr('y', 65)
    .attr('text-anchor', 'middle')
    .style('font-size', '14px')
    .style('fill', '#999')
    .text(`Target: ${d3.format(',.0f')(targetValue)}`);

  // 10. Add dimension label at top (if present)
  if (dimension) {
    svg.append('text')
      .attr('x', width / 2)
      .attr('y', 30)
      .attr('text-anchor', 'middle')
      .style('font-size', '16px')
      .style('font-weight', '600')
      .style('fill', '#333')
      .text(dimension);
  }

  const selectionLayer = svg.append('g');
  const hoveringLayer = svg.append('g');

  return { hoveringLayer, viz: svg.node() };
}

function formatValue(value) {
  if (value === 0) return '0.0M';
  if (value >= 1000000) {
    return (value / 1000000).toFixed(1) + 'M';
  } else if (value >= 1000) {
    return (value / 1000).toFixed(0) + 'K';
  }
  return d3.format(',.0f')(value);
}

async function renderViz (rawData, encodingMap, selectedMarksIds, styles) {
  const encodedData = getEncodedData(rawData, encodingMap);
  const content = document.getElementById('content');
  content.innerHTML = '';
  const chart = await DonutGaugeChart(encodedData, encodingMap, content.offsetWidth, content.offsetHeight, selectedMarksIds, styles);
  content.appendChild(chart.viz);
  return chart;
}

async function getEncodingMap () {
  const worksheet = tableau.extensions.worksheetContent.worksheet;
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

window.onload = function() {
  tableau.extensions.initializeAsync().then(async () => {
    const worksheet = tableau.extensions.worksheetContent.worksheet;
    let summaryData = {}, encodingMap = {}, selectedMarks = new Map();
    const styles = tableau.extensions.environment.workbookFormatting?.formattingSheets?.find(x => x.classNameKey === 'tableau-worksheet')?.cssProperties;

    const updateDataAndRender = async () => {
      [summaryData, encodingMap] = await Promise.all([getSummaryDataTable(worksheet), getEncodingMap(worksheet)]);
      selectedMarks = await getSelection(worksheet, summaryData);
      await renderViz(summaryData, encodingMap, selectedMarks, styles);
    };

    onresize = async () => {
      await renderViz(summaryData, encodingMap, selectedMarks, styles);
    };
    
    worksheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, updateDataAndRender);
    
    updateDataAndRender();
  }).catch(err => {
    console.error('Extension initialization failed:', err);
    document.getElementById('content').innerHTML = '<div style="padding:20px;color:red;">Error: ' + err.message + '</div>';
  });
};

async function getSelection (worksheet, allMarks) {
  try {
    const selectedMarks = await worksheet.getSelectedMarksAsync();
    return findIdsOfSelectedMarks(allMarks, selectedMarks);
  } catch (e) {
    return new Map();
  }
}

function findIdsOfSelectedMarks (allMarks, selectedMarks) {
  const columns = selectedMarks.data[0].columns;
  const selectedMarkMap = new Map();
  const selectedMarksIds = new Map();
  
  for (const selectedMark of convertToListOfNamedRows(selectedMarks.data[0])) {
    let key = '';
    for (const col of columns) {
      key += selectedMark[col.fieldName].value + '\x00';
    }
    selectedMarkMap.set(key, selectedMark);
  }
  
  let tupleId = 1;
  for (const mark of allMarks) {
    let key = '';
    for (const col of columns) {
      key += mark[col.fieldName].value + '\x00';
    }
    if (selectedMarkMap.has(key)) {
      selectedMarksIds.set(tupleId);
    }
    tupleId++;
  }
  return selectedMarksIds;
}

function convertToListOfNamedRows (dataTablePage) {
  const rows = [];
  const columns = dataTablePage.columns;
  const data = dataTablePage.data;
  
  for (let i = 0; i < data.length; ++i) {
    const row = {};
    for (let j = 0; j < columns.length; ++j) {
      row[columns[j].fieldName] = data[i][columns[j].index];
    }
    row.tupleId = i + 1;
    rows.push(row);
  }
  return rows;
}

async function getSummaryDataTable (worksheet) {
  let rows = [];
  const dataTableReader = await worksheet.getSummaryDataReaderAsync(undefined, { ignoreSelection: true });
  for (let currentPage = 0; currentPage < dataTableReader.pageCount; currentPage++) {
    const dataTablePage = await dataTableReader.getPageAsync(currentPage);
    rows = rows.concat(convertToListOfNamedRows(dataTablePage));
  }
  await dataTableReader.releaseAsync();
  return rows;
}

function getEncodedData (data, encodingMap) {
  const encodedData = [];
  let tupleId = 1;
  
  for (const row of data) {
    const encodedRow = {};
    for (const encName in encodingMap) {
      const fields = encodingMap[encName];
      encodedRow[encName] = [];
      for (const field of fields) {
        encodedRow[encName].push(row[field.name]);
      }
    }
    encodedRow.tupleId = tupleId++;
    encodedData.push(encodedRow);
  }
  return encodedData;
}

function getWorksheet () {
  return tableau.extensions.worksheetContent ? tableau.extensions.worksheetContent.worksheet : tableau.extensions.dashboardContent.dashboard.worksheets[0];
}

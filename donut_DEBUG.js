/* global d3 */
/* global tinycolor */
/* global tableau */

const backgroundColor = tinycolor('white');
const palette = ['#4e79a7', '#f28e2c', '#e15759', '#76b7b2', '#59a14f', '#edc949', '#af7aa1', '#ff9da7', '#9c755f', '#bab0ab'];

async function DonutChart (encodedData, encodingMap, width, height, selectedTupleIds, styles) {
  const margin = { top: 40, right: 20, bottom: 40, left: 20 };
  const innerWidth = width - margin.left - margin.right;
  const innerHeight = height - margin.top - margin.bottom;
  const radius = Math.min(innerWidth, innerHeight) / 2;

  console.log('=== DONUT CHART DEBUG ===');
  console.log('Encoded data rows:', encodedData.length);
  console.log('Encoding map:', encodingMap);
  console.log('First row sample:', encodedData[0]);

  // 1. Aggregate Data
  const aggregatedData = [];
  const dataMap = new Map();

  encodedData.forEach(row => {
    let cat = 'Unknown';
    let val = 0;

    // Try to get category from encoding
    if (row.category && row.category.length > 0 && row.category[0]) {
      cat = row.category[0].formattedValue || row.category[0].value || 'Unknown';
    }

    // Try to get measure from encoding
    if (row.measure && row.measure.length > 0 && row.measure[0]) {
      val = parseFloat(row.measure[0].value) || 0;
    }

    console.log('Row data - Category:', cat, 'Value:', val);

    if (dataMap.has(cat)) {
      dataMap.set(cat, dataMap.get(cat) + val);
    } else {
      dataMap.set(cat, val);
    }
  });

  let tupleIdCounter = 1;
  dataMap.forEach((value, key) => {
    aggregatedData.push({ category: key, value: value, tupleId: tupleIdCounter++ });
  });

  // Sort by value (descending)
  aggregatedData.sort((a, b) => b.value - a.value);

  console.log('Aggregated data:', aggregatedData);
  console.log('Total categories:', aggregatedData.length);

  // If no data, show error message
  if (aggregatedData.length === 0 || d3.sum(aggregatedData, d => d.value) === 0) {
    const errorSvg = d3.create('svg')
      .attr('width', width)
      .attr('height', height);
    
    errorSvg.append('text')
      .attr('x', width / 2)
      .attr('y', height / 2)
      .attr('text-anchor', 'middle')
      .style('font-size', '16px')
      .style('fill', '#999')
      .text('No data available. Check dimension and measure fields.');

    return { viz: errorSvg.node() };
  }

  // 2. Create SVG
  const svg = d3.create('svg')
    .attr('class', tableau.ClassNameKey.Worksheet)
    .attr('width', width)
    .attr('height', height)
    .attr('viewBox', [0, 0, width, height])
    .attr('style', 'max-width: 100%; height: auto;')
    .attr('font-family', styles?.fontFamily || 'Arial')
    .attr('font-size', styles?.fontSize || '12px');

  const chartGroup = svg.append('g')
    .attr('transform', `translate(${width / 2},${height / 2})`);

  // 3. Create pie and arc generators
  const pie = d3.pie()
    .value(d => d.value)
    .sort(null);

  const arc = d3.arc()
    .innerRadius(radius * 0.5)
    .outerRadius(radius);

  const arcHover = d3.arc()
    .innerRadius(radius * 0.5)
    .outerRadius(radius * 1.05);

  const colorScale = d3.scaleOrdinal()
    .domain(aggregatedData.map((d, i) => i))
    .range(palette);

  // 4. Draw donut slices
  const slices = chartGroup.selectAll('.slice')
    .data(pie(aggregatedData))
    .enter()
    .append('path')
    .attr('class', 'slice')
    .attr('d', arc)
    .attr('fill', (d, i) => getColor(colorScale(i), selectedTupleIds))
    .attr('stroke', 'white')
    .attr('stroke-width', 2);

  // 5. Add percentage labels
  const labelArc = d3.arc()
    .innerRadius(radius * 0.75)
    .outerRadius(radius * 0.75);

  chartGroup.selectAll('.label')
    .data(pie(aggregatedData))
    .enter()
    .append('text')
    .attr('class', 'label')
    .attr('transform', d => `translate(${labelArc.centroid(d)})`)
    .attr('text-anchor', 'middle')
    .style('font-size', '12px')
    .style('font-weight', 'bold')
    .style('fill', 'white')
    .style('text-shadow', '1px 1px 2px rgba(0,0,0,0.8)')
    .text(d => {
      const percentage = ((d.data.value / d3.sum(aggregatedData, d => d.value)) * 100).toFixed(1);
      return percentage > 5 ? `${percentage}%` : '';
    });

  // 6. Add center total
  const totalValue = d3.sum(aggregatedData, d => d.value);
  
  chartGroup.append('text')
    .attr('text-anchor', 'middle')
    .attr('dy', '-0.5em')
    .style('font-size', '14px')
    .style('font-weight', '600')
    .style('fill', '#666')
    .text('Total');

  chartGroup.append('text')
    .attr('text-anchor', 'middle')
    .attr('dy', '1em')
    .style('font-size', '28px')
    .style('font-weight', 'bold')
    .style('fill', '#333')
    .text(d3.format(',.0f')(totalValue));

  // 7. Add legend
  const legend = svg.append('g')
    .attr('class', 'legend')
    .attr('transform', `translate(${width - 150}, 20)`);

  const legendItems = legend.selectAll('.legend-item')
    .data(aggregatedData)
    .enter()
    .append('g')
    .attr('class', 'legend-item')
    .attr('transform', (d, i) => `translate(0, ${i * 20})`);

  legendItems.append('rect')
    .attr('width', 15)
    .attr('height', 15)
    .attr('fill', (d, i) => colorScale(i));

  legendItems.append('text')
    .attr('x', 20)
    .attr('y', 12)
    .style('font-size', '11px')
    .text(d => {
      const label = d.category.length > 15 ? d.category.substring(0, 15) + '...' : d.category;
      return `${label}: ${d3.format(',.0f')(d.value)}`;
    });

  // 8. Interaction Layers
  const selectionLayer = svg.append('g');
  const hoveringLayer = svg.append('g');
  const slicesPerTupleId = getSlicesPerTupleId(slices);
  renderSelection(selectedTupleIds, slicesPerTupleId, selectionLayer, hoveringLayer, colorScale, chartGroup, arc);

  console.log('=== CHART RENDERED SUCCESSFULLY ===');

  return { hoveringLayer, slicesPerTupleId, arc, arcHover, chartGroup, viz: svg.node() };
}

function getSlicesPerTupleId (slices) {
  const map = new Map();
  slices.each(function (d) { 
    map.set(d.data.tupleId, { 
      element: d3.select(this), 
      data: d 
    }); 
  });
  return map;
}

function renderSelection (selectedTupleIds, slicesPerTupleId, selectionLayer, highlightingLayer, colorScale, chartGroup, arc) {
  selectionLayer.selectAll('*').remove();
  highlightingLayer.selectAll('*').remove();
  
  const selectedSlices = [];
  for (const id of selectedTupleIds.keys()) {
    const slice = slicesPerTupleId.get(id);
    if (slice) selectedSlices.push(slice);
  }

  selectedSlices.forEach(slice => {
    chartGroup.append('path')
      .datum(slice.data)
      .attr('class', 'selected-slice')
      .attr('d', arc)
      .attr('fill', 'none')
      .attr('stroke', 'black')
      .attr('stroke-width', 3)
      .attr('transform', chartGroup.attr('transform'));
  });
}

function renderHoveredElements (hoveredTupleIds, slicesPerTupleId, hoveringLayer, arcHover, chartGroup) {
  if (!hoveringLayer) return;
  hoveringLayer.selectAll('*').remove();
  
  const hoveredSlices = [];
  for (const id of hoveredTupleIds.keys()) {
    const slice = slicesPerTupleId.get(id);
    if (slice) hoveredSlices.push(slice);
  }

  hoveredSlices.forEach(slice => {
    slice.element.transition()
      .duration(200)
      .attr('d', arcHover);
  });
}

function clearHoveredSlices(slicesPerTupleId, arc) {
  slicesPerTupleId.forEach(slice => {
    slice.element.transition()
      .duration(200)
      .attr('d', arc);
  });
}

async function renderViz (rawData, encodingMap, selectedMarksIds, styles) {
  const encodedData = getEncodedData(rawData, encodingMap);
  console.log('Rendering viz with encoded data:', encodedData.length, 'rows');
  
  const content = document.getElementById('content');
  content.innerHTML = '';
  
  const donutChart = await DonutChart(encodedData, encodingMap, content.offsetWidth, content.offsetHeight, selectedMarksIds, styles);
  content.appendChild(donutChart.viz);
  
  return donutChart;
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
  console.log('Encoding map created:', encodingMap);
  return encodingMap;
}

window.onload = function() {
  tableau.extensions.initializeAsync().then(async () => {
    console.log('✓ Extension initialized');
    
    const worksheet = tableau.extensions.worksheetContent.worksheet;
    let summaryData = {}, encodingMap = {}, selectedMarks = new Map(), hoveredMarks = new Map(), slicesPerTupleId = new Map(), hoveringLayer, arc, arcHover, chartGroup;
    const styles = tableau.extensions.environment.workbookFormatting?.formattingSheets?.find(x => x.classNameKey === 'tableau-worksheet')?.cssProperties;

    const updateDataAndRender = async () => {
      try {
        [summaryData, encodingMap] = await Promise.all([getSummaryDataTable(worksheet), getEncodingMap(worksheet)]);
        console.log('Summary data rows:', summaryData.length);
        
        selectedMarks = await getSelection(worksheet, summaryData);
        const result = await renderViz(summaryData, encodingMap, selectedMarks, styles);
        
        hoveringLayer = result.hoveringLayer;
        slicesPerTupleId = result.slicesPerTupleId;
        arc = result.arc;
        arcHover = result.arcHover;
        chartGroup = result.chartGroup;
      } catch (error) {
        console.error('Error in updateDataAndRender:', error);
        document.getElementById('content').innerHTML = '<div style="padding:40px;color:red;">Error: ' + error.message + '</div>';
      }
    };

    onresize = async () => {
      const result = await renderViz(summaryData, encodingMap, selectedMarks, styles);
      hoveringLayer = result.hoveringLayer;
      slicesPerTupleId = result.slicesPerTupleId;
      arc = result.arc;
      arcHover = result.arcHover;
      chartGroup = result.chartGroup;
    };
    
    worksheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, updateDataAndRender);
    
    document.body.addEventListener('click', async (e) => {
      onClick(e, selectedMarks, hoveredMarks);
      const result = await renderViz(summaryData, encodingMap, selectedMarks, styles);
      hoveringLayer = result.hoveringLayer;
      slicesPerTupleId = result.slicesPerTupleId;
      arc = result.arc;
      arcHover = result.arcHover;
      chartGroup = result.chartGroup;
    });
    
    document.body.addEventListener('mousemove', e => onMouseMove(e, hoveredMarks, slicesPerTupleId, hoveringLayer, arcHover, chartGroup, arc));
    document.body.addEventListener('mouseout', e => {
      clearHoveredMarks(hoveredMarks);
      clearHoveredSlices(slicesPerTupleId, arc);
    });
    
    updateDataAndRender();
  }).catch(err => {
    console.error('Extension initialization failed:', err);
    document.getElementById('content').innerHTML = '<div style="padding:20px;color:red;">Error: ' + err.message + '</div>';
  });
};

function onClick (e, selectedTupleIds, hoveredTupleIds) {
  const elem = d3.select(document.elementFromPoint(e.pageX, e.pageY));
  const data = elem?.datum();
  const tupleId = data?.data?.tupleId;
  
  if (elem && tupleId !== null && tupleId !== undefined) {
    if (selectedTupleIds.has(tupleId)) {
      if (selectedTupleIds.size === 1) selectedTupleIds.clear();
      else if (e.ctrlKey) selectedTupleIds.delete(tupleId);
      else {
        selectedTupleIds.clear();
        selectedTupleIds.set(tupleId);
      }
    } else {
      if (!e.ctrlKey) selectedTupleIds.clear();
      selectedTupleIds.set(tupleId);
    }
  } else if (!e.ctrlKey) {
    selectedTupleIds.clear();
  }
  
  selectTuples(e.pageX, e.pageY, selectedTupleIds, hoveredTupleIds);
}

async function selectTuples (x, y, selectedTupleIds, hoveredTupleIds) {
  clearHoveredMarks(hoveredTupleIds);
  getWorksheet().selectTuplesAsync([...selectedTupleIds.keys()], tableau.SelectOptions.Simple, { tooltipAnchorPoint: { x, y } });
}

async function onMouseMove (e, hoveredTupleIds, slicesPerTupleId, hoveringLayer, arcHover, chartGroup, arc) {
  const elem = d3.select(document.elementFromPoint(e.pageX, e.pageY));
  const data = elem?.node() ? elem.datum() : undefined;
  const tupleId = data?.data?.tupleId;
  
  const hadHoveredTupleBefore = hoveredTupleIds.size !== 0;
  clearHoveredMarks(hoveredTupleIds);
  clearHoveredSlices(slicesPerTupleId, arc);
  
  if (elem && tupleId !== null && tupleId !== undefined) {
    hoveredTupleIds.set(tupleId);
    getWorksheet().hoverTupleAsync(parseInt(tupleId), { tooltipAnchorPoint: { x: e.pageX, y: e.pageY } });
  } else if (hadHoveredTupleBefore) {
    getWorksheet().hoverTupleAsync(parseInt(tupleId), { tooltipAnchorPoint: { x: e.pageX, y: e.pageY } });
  }
  
  renderHoveredElements(hoveredTupleIds, slicesPerTupleId, hoveringLayer, arcHover, chartGroup);
}

function clearHoveredMarks (hoveredTupleIds) {
  hoveredTupleIds.clear();
}

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
  console.log('Got summary data:', rows.length, 'rows');
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
  
  console.log('Encoded', encodedData.length, 'rows');
  return encodedData;
}

function getColor (color, selectedTupleIds) {
  return (selectedTupleIds.size > 0) ? calculateFogColor(color) : color;
}

const fogBlendFactor = getFogBlendFactor(backgroundColor);
const { foggedBackgroundRed, foggedBackgroundGreen, foggedBackgroundBlue } = computeFoggedBackgroundColor(backgroundColor, fogBlendFactor);

function computeFoggedBackgroundColor (color, fogBlendFactor) {
  const CloseToWhite = 245;
  if (color.r >= CloseToWhite && color.g >= CloseToWhite && color.b >= CloseToWhite) {
    color = tinycolor({ r: CloseToWhite, g: CloseToWhite, b: CloseToWhite });
  }
  color = color.toRgb();
  return {
    foggedBackgroundRed: (1 - fogBlendFactor) * color.r >>> 0,
    foggedBackgroundGreen: (1 - fogBlendFactor) * color.g >>> 0,
    foggedBackgroundBlue: (1 - fogBlendFactor) * color.b >>> 0
  };
}

function calculateFogColor (colorStr) {
  const color = tinycolor(colorStr).toRgb();
  return tinycolor({
    r: foggedBackgroundRed + color.r * fogBlendFactor >>> 0,
    g: foggedBackgroundGreen + color.g * fogBlendFactor >>> 0,
    b: foggedBackgroundBlue + color.b * fogBlendFactor >>> 0
  }).toHexString();
}

function getFogBlendFactor (color) {
  color = color.toRgb();
  const DefaultFogBlendFactor = 0.1850000023841858;
  const DarkBgFogBlendFactor = 0.2750000059604645;
  const DarkBgThreshold = 75;
  return (color.r <= DarkBgThreshold && color.g <= DarkBgThreshold && color.b <= DarkBgThreshold) ? DarkBgFogBlendFactor : DefaultFogBlendFactor;
}

function getWorksheet () {
  return tableau.extensions.worksheetContent ? tableau.extensions.worksheetContent.worksheet : tableau.extensions.dashboardContent.dashboard.worksheets[0];
}

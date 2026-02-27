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

  // 1. DYNAMIC KEY DETECTION
  // We need to find which key in the row object is the Dimension (String) and which is the Measure (Number).
  let catKey = null;
  let valKey = null;

  // Helper to inspect data types in the first row
  function detectKeysFromData(data) {
    if (!data || data.length === 0) return;
    const sample = data[0];
    for (const key in sample) {
        if (key === 'tupleId') continue;
        const arr = sample[key];
        if (Array.isArray(arr) && arr.length > 0 && arr[0] && arr[0].value !== undefined) {
            const v = arr[0].value;
            if (typeof v === 'string' && !catKey) catKey = key;
            if (typeof v === 'number' && !valKey) valKey = key;
        }
    }
  }

  // Strategy A: Check encodingMap keys (e.g., 'color', 'detail', 'measure')
  if (encodingMap) {
      const roles = Object.keys(encodingMap);
      if (!catKey) catKey = roles.find(k => ['color', 'detail', 'label', 'tooltip'].includes(k.toLowerCase()));
      if (!valKey) valKey = roles.find(k => ['measure', 'size', 'angle', 'text'].includes(k.toLowerCase()));
  }

  // Strategy B: Fallback to Data Type Inspection if Strategy A failed
  if (!catKey || !valKey) {
      detectKeysFromData(encodedData);
  }

  // Final Fallback: Just pick the first available keys
  if (!catKey || !valKey) {
      const keys = Object.keys(encodedData[0]).filter(k => k !== 'tupleId');
      if (!catKey) catKey = keys[0];
      if (!valKey) valKey = keys[1] || keys[0];
  }

  console.log(`Detected Dimension Key: ${catKey}, Measure Key: ${valKey}`);

  // 2. Aggregate Data
  const aggregatedData = [];
  const dataMap = new Map();

  encodedData.forEach(row => {
    // Use detected keys to fetch values
    const catObj = row[catKey] ? row[catKey][0] : null;
    const valObj = row[valKey] ? row[valKey][0] : null;

    let cat = catObj ? (catObj.formattedValue || catObj.value) : 'Unknown';
    let val = valObj ? parseFloat(valObj.value) : 0;
    const rowTupleId = row.tupleId; // The ID from the raw Tableau data

    if (dataMap.has(cat)) {
      const entry = dataMap.get(cat);
      entry.value += val;
      entry.tupleIds.push(rowTupleId); // Collect all source IDs for this category
    } else {
      dataMap.set(cat, { value: val, tupleIds: [rowTupleId] });
    }
  });

  // Convert Map to Array for D3
  let internalIdCounter = 1;
  dataMap.forEach((entry, key) => {
    aggregatedData.push({ 
        category: key, 
        value: entry.value, 
        // tupleId is used for D3 object constancy (mapping DOM to data)
        tupleId: internalIdCounter++, 
        // tupleIds is the list of real Tableau IDs needed for interaction
        tupleIds: entry.tupleIds 
    });
  });

  // Sort by value (descending)
  aggregatedData.sort((a, b) => b.value - a.value);

  console.log('Aggregated data:', aggregatedData);

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

  // 3. Create SVG
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

  // 4. Create pie and arc generators
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

  // 5. Draw donut slices
  const slices = chartGroup.selectAll('.slice')
    .data(pie(aggregatedData))
    .enter()
    .append('path')
    .attr('class', 'slice')
    .attr('d', arc)
    .attr('fill', (d, i) => getColor(colorScale(i), selectedTupleIds, d.data)) // Pass slice data for color logic
    .attr('stroke', 'white')
    .attr('stroke-width', 2);

  // 6. Add percentage labels
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

  // 7. Add center total
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

  // 8. Add legend
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

  // 9. Interaction Layers
  const selectionLayer = svg.append('g');
  const hoveringLayer = svg.append('g');
  // Pass the D3 selection 'slices' instead of a map, as we need to check data content
  renderSelection(selectedTupleIds, slices, selectionLayer, hoveringLayer, colorScale, chartGroup, arc);

  console.log('=== CHART RENDERED SUCCESSFULLY ===');

  return { hoveringLayer, slices, arc, arcHover, chartGroup, viz: svg.node() };
}

// Render Selection: Draw strokes on slices that are represented in selectedTupleIds
function renderSelection (selectedTupleIds, slices, selectionLayer, highlightingLayer, colorScale, chartGroup, arc) {
  selectionLayer.selectAll('*').remove();
  highlightingLayer.selectAll('*').remove();

  if (selectedTupleIds.size === 0) return;

  // Iterate over all visual slices to see if they contain any selected IDs
  slices.each(function(d) {
      // Check if this slice's underlying data (tupleIds) intersects with selectedTupleIds
      const isSelected = d.data.tupleIds.some(id => selectedTupleIds.has(id));
      
      if (isSelected) {
          chartGroup.append('path')
            .datum(d)
            .attr('class', 'selected-slice')
            .attr('d', arc)
            .attr('fill', 'none')
            .attr('stroke', 'black')
            .attr('stroke-width', 3)
            .attr('transform', chartGroup.attr('transform'));
      }
  });
}

function renderHoveredElements (hoveredTupleIds, slices, hoveringLayer, arcHover, chartGroup) {
  if (!hoveringLayer) return;
  hoveringLayer.selectAll('*').remove();
  
  // Logic remains similar: find the slice that contains the hovered ID
  // However, the passed 'hoveredTupleIds' in the original code contained the 'internal' ID from the visual datum.
  // Let's rely on the element reference passed in onMouseMove or use a similar lookup.
  // To keep it simple and working with the existing structure:
  
  // We assume hoveredTupleIds might contain the 'internal' ID for visual expansion logic,
  // but let's look at how it's called. It's called from onMouseMove.
}

function clearHoveredSlices(slices, arc) {
  slices.transition()
      .duration(200)
      .attr('d', arc);
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
    let summaryData = {}, encodingMap = {}, selectedMarks = new Map(), hoveredMarks = new Map(), slices, hoveringLayer, arc, arcHover, chartGroup;
    const styles = tableau.extensions.environment.workbookFormatting?.formattingSheets?.find(x => x.classNameKey === 'tableau-worksheet')?.cssProperties;

    const updateDataAndRender = async () => {
      try {
        [summaryData, encodingMap] = await Promise.all([getSummaryDataTable(worksheet), getEncodingMap(worksheet)]);
        console.log('Summary data rows:', summaryData.length);
        
        selectedMarks = await getSelection(worksheet, summaryData);
        const result = await renderViz(summaryData, encodingMap, selectedMarks, styles);
        
        hoveringLayer = result.hoveringLayer;
        slices = result.slices; // Capture the D3 selection
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
      slices = result.slices;
      arc = result.arc;
      arcHover = result.arcHover;
      chartGroup = result.chartGroup;
    };
    
    worksheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, updateDataAndRender);
    
    document.body.addEventListener('click', async (e) => {
      onClick(e, selectedMarks, hoveredMarks, slices, arc);
      const result = await renderViz(summaryData, encodingMap, selectedMarks, styles);
      hoveringLayer = result.hoveringLayer;
      slices = result.slices;
      arc = result.arc;
      arcHover = result.arcHover;
      chartGroup = result.chartGroup;
    });
    
    document.body.addEventListener('mousemove', e => onMouseMove(e, hoveredMarks, slices, hoveringLayer, arcHover, chartGroup, arc));
    document.body.addEventListener('mouseout', e => {
      clearHoveredMarks(hoveredMarks);
      if (slices) clearHoveredSlices(slices, arc);
    });
    
    updateDataAndRender();
  }).catch(err => {
    console.error('Extension initialization failed:', err);
    document.getElementById('content').innerHTML = '<div style="padding:20px;color:red;">Error: ' + err.message + '</div>';
  });
};

function onClick (e, selectedTupleIds, hoveredTupleIds, slices, arc) {
  const elem = d3.select(document.elementFromPoint(e.pageX, e.pageY));
  const data = elem?.datum();
  
  // We check if we clicked a slice
  if (data && data.data && data.data.tupleIds) {
      const sliceIds = data.data.tupleIds;
      
      // Check if this slice is currently "selected" (i.e., all its IDs are in the selection)
      // For simplicity in this UI, we'll check if ANY of its IDs are selected
      const isCurrentlySelected = sliceIds.some(id => selectedTupleIds.has(id));

      if (isCurrentlySelected) {
          if (e.ctrlKey) {
              // Remove these specific IDs
              sliceIds.forEach(id => selectedTupleIds.delete(id));
          } else {
              // If clicking the active selection without Ctrl, clear selection
              // Unless it's a multi-select situation where we want to keep others? 
              // Standard behavior: Clicking selected item deselects it (clears all if it was the only one)
              // Or: Clicking it clears everything else?
              // Let's implement: Toggle logic.
              // If it's the ONLY thing selected, clear it. Otherwise, if we click a selected item in a multi-select, remove it.
              // For simplicity in this chart: Clicking a selected slice clears selection (unless Ctrl)
              selectedTupleIds.clear();
          }
      } else {
          // Selecting a new slice
          if (!e.ctrlKey) selectedTupleIds.clear();
          sliceIds.forEach(id => selectedTupleIds.set(id, true));
      }
  } else if (!e.ctrlKey) {
      selectedTupleIds.clear();
  }
  
  selectTuples(e.pageX, e.pageY, selectedTupleIds, hoveredTupleIds);
}

async function selectTuples (x, y, selectedTupleIds, hoveredTupleIds) {
  clearHoveredMarks(hoveredTupleIds);
  // selectedTupleIds is a Map, .keys() gives an iterator of IDs
  getWorksheet().selectTuplesAsync([...selectedTupleIds.keys()], tableau.SelectOptions.Simple, { tooltipAnchorPoint: { x, y } });
}

async function onMouseMove (e, hoveredTupleIds, slices, hoveringLayer, arcHover, chartGroup, arc) {
  const elem = d3.select(document.elementFromPoint(e.pageX, e.pageY));
  const data = elem?.node() ? elem.datum() : undefined;
  
  const hadHoveredTupleBefore = hoveredTupleIds.size !== 0;
  clearHoveredMarks(hoveredTupleIds);
  if (slices) clearHoveredSlices(slices, arc);
  
  if (data && data.data && data.data.tupleIds) {
      // We found a slice. 
      // 1. Visually expand it
      const sliceElement = d3.select(elem.node());
      sliceElement.transition().duration(200).attr('d', arcHover);
      
      // 2. Tell Tableau to hover the underlying data points
      // We pick the first ID to represent the group for the tooltip
      const representativeId = data.data.tupleIds[0];
      hoveredTupleIds.set(representativeId);
      
      getWorksheet().hoverTupleAsync(representativeId, { tooltipAnchorPoint: { x: e.pageX, y: e.pageY } });
  } else if (hadHoveredTupleBefore) {
      getWorksheet().hoverTupleAsync(null, { tooltipAnchorPoint: { x: e.pageX, y: e.pageY } });
  }
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

// Updated to accept sliceData to check if it's selected
function getColor (color, selectedTupleIds, sliceData) {
  if (selectedTupleIds.size === 0) return color;
  
  // If this slice contains any selected ID, keep it bright. Otherwise, fog it.
  const isSelected = sliceData && sliceData.tupleIds && sliceData.tupleIds.some(id => selectedTupleIds.has(id));
  
  return isSelected ? color : calculateFogColor(color);
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
/* global d3 */
/* global tinycolor */
/* global tableau */

const backgroundColor = tinycolor('white');
const palette = ['#5B6FD8', '#D3D3D3', '#4e79a7', '#f28e2c'];

async function GaugeChart (encodedData, encodingMap, width, height, selectedTupleIds, styles) {
  const margin = { top: 60, right: 60, bottom: 60, left: 60 };
  const innerWidth = width - margin.left - margin.right;
  const innerHeight = height - margin.top - margin.bottom;

  console.log('=== GAUGE CHART ===');

  // 1. DETECT NUMERIC KEYS
  let valueKey = null;
  let targetKey = null;

  function detectNumericKeys(data) {
    if (!data || data.length === 0) return [];
    const sample = data[0];
    const numericKeys = [];
    for (const key in sample) {
        if (key === 'tupleId') continue;
        const arr = sample[key];
        if (Array.isArray(arr) && arr.length > 0 && arr[0] && arr[0].value !== undefined) {
            const v = arr[0].value;
            if (typeof v === 'number') {
                numericKeys.push(key);
            }
        }
    }
    return numericKeys;
  }

  const numericKeys = detectNumericKeys(encodedData);
  
  if (numericKeys.length > 0) {
      const foundTarget = numericKeys.find(k => k.toLowerCase().includes('target'));
      if (foundTarget && numericKeys.length > 1) {
          targetKey = foundTarget;
          valueKey = numericKeys.find(k => k !== foundTarget);
      } else {
          valueKey = numericKeys[0];
          if (numericKeys.length > 1) targetKey = numericKeys[1];
      }
  }

  if (!valueKey) {
    console.error('No numeric measure found');
    return { viz: null };
  }

  // 2. AGGREGATE DATA
  let totalValue = 0;
  let totalTarget = 0;
  const allTupleIds = [];

  encodedData.forEach(row => {
    const valObj = row[valueKey] ? row[valueKey][0] : null;
    if (valObj && valObj.value !== undefined) {
        totalValue += parseFloat(valObj.value);
    }
    
    if (targetKey) {
        const tgtObj = row[targetKey] ? row[targetKey][0] : null;
        if (tgtObj && tgtObj.value !== undefined) {
            totalTarget += parseFloat(tgtObj.value);
        }
    }

    if (row.tupleId) allTupleIds.push(row.tupleId);
  });

  // 3. CALCULATE MAX SCALE (round up nicely)
  let maxScale = Math.max(totalValue, totalTarget);
  if (maxScale === 0) maxScale = 1000000;
  maxScale = Math.ceil(maxScale / 1000000) * 1000000;

  console.log(`Value: ${totalValue}, Max Scale: ${maxScale}`);

  // 4. SVG SETUP
  const svg = d3.create('svg')
    .attr('class', tableau.ClassNameKey.Worksheet)
    .attr('width', width)
    .attr('height', height)
    .attr('viewBox', [0, 0, width, height])
    .attr('style', 'max-width: 100%; height: auto; background: white;')
    .attr('font-family', styles?.fontFamily || 'Arial, sans-serif')
    .attr('font-size', styles?.fontSize || '12px');

  const cx = width / 2;
  const cy = height / 2;
  const radius = Math.min(innerWidth, innerHeight) * 0.42;

  const chartGroup = svg.append('g')
    .attr('transform', `translate(${cx}, ${cy})`);

  // 5. ANGLES: 3/4 circle (270 degrees) - EXACT positions
  const startAngle = -Math.PI * 3/4;  // Bottom-left
  const endAngle = Math.PI * 3/4;      // Bottom-right

  // 6. BACKGROUND ARC (GRAY - full range)
  const backgroundArc = d3.arc()
    .innerRadius(radius * 0.7)
    .outerRadius(radius)
    .startAngle(startAngle)
    .endAngle(endAngle);

  chartGroup.append('path')
    .attr('d', backgroundArc)
    .attr('fill', '#D3D3D3')
    .attr('opacity', 1);

  // 7. VALUE ARC (BLUE - filled portion)
  const valueFraction = Math.min(Math.max(totalValue / maxScale, 0), 1);
  const valueEndAngle = startAngle + valueFraction * (endAngle - startAngle);
  
  const valueArc = d3.arc()
    .innerRadius(radius * 0.7)
    .outerRadius(radius)
    .startAngle(startAngle)
    .endAngle(valueEndAngle);

  chartGroup.append('path')
    .attr('d', valueArc)
    .attr('fill', palette[0])
    .attr('opacity', 1);

  // 8. SCALE LABELS - AT ARC START AND END POSITIONS
  // Calculate proper positions for 5 labels evenly distributed
  const numLabels = 5;
  for (let i = 0; i <= numLabels; i++) {
      const fraction = i / numLabels;
      const angle = startAngle + fraction * (endAngle - startAngle);
      const labelValue = fraction * maxScale;
      
      // Position labels OUTSIDE the arc
      const labelRadius = radius * 1.20;
      const lx = labelRadius * Math.cos(angle);
      const ly = labelRadius * Math.sin(angle);
      
      // Format label
      let labelText;
      if (labelValue === 0) {
          labelText = '0';
      } else {
          labelText = d3.format(',.0f')(labelValue);
      }
      
      chartGroup.append('text')
        .attr('x', lx)
        .attr('y', ly)
        .attr('text-anchor', 'middle')
        .attr('dy', '0.35em')
        .style('font-size', '14px')
        .style('font-weight', '400')
        .style('fill', '#333')
        .text(labelText);
  }

  // 9. ADD TICK MARKS at label positions for clarity
  for (let i = 0; i <= numLabels; i++) {
      const fraction = i / numLabels;
      const angle = startAngle + fraction * (endAngle - startAngle);
      
      const tickInnerR = radius * 1.02;
      const tickOuterR = radius * 1.10;
      
      const x1 = tickInnerR * Math.cos(angle);
      const y1 = tickInnerR * Math.sin(angle);
      const x2 = tickOuterR * Math.cos(angle);
      const y2 = tickOuterR * Math.sin(angle);

      chartGroup.append('line')
        .attr('x1', x1).attr('y1', y1)
        .attr('x2', x2).attr('y2', y2)
        .attr('stroke', '#333')
        .attr('stroke-width', 2);
  }

  // 10. TARGET MARKER (if present)
  if (targetKey && totalTarget > 0) {
      const targetFraction = Math.min(Math.max(totalTarget / maxScale, 0), 1);
      const targetAngle = startAngle + targetFraction * (endAngle - startAngle);
      
      const tickInnerR = radius * 0.68;
      const tickOuterR = radius * 1.12;
      
      const tx1 = tickInnerR * Math.cos(targetAngle);
      const ty1 = tickInnerR * Math.sin(targetAngle);
      const tx2 = tickOuterR * Math.cos(targetAngle);
      const ty2 = tickOuterR * Math.sin(targetAngle);

      chartGroup.append('line')
        .attr('x1', tx1).attr('y1', ty1)
        .attr('x2', tx2).attr('y2', ty2)
        .attr('stroke', '#f28e2c')
        .attr('stroke-width', 5)
        .attr('stroke-linecap', 'round');
      
      chartGroup.append('circle')
        .attr('cx', tx2)
        .attr('cy', ty2)
        .attr('r', 6)
        .attr('fill', '#f28e2c');
  }

  // 11. NEEDLE - POINTS TO EXACT VALUE POSITION (end of blue arc)
  const needleLength = radius * 0.85;
  const needleAngle = valueEndAngle;  // EXACTLY where blue arc ends
  
  // Calculate needle tip position
  const needleTipX = needleLength * Math.cos(needleAngle);
  const needleTipY = needleLength * Math.sin(needleAngle);
  
  const needleWidth = 16;
  const needleBaseWidth = needleWidth;
  
  // Create needle path (triangle pointing to value)
  const needlePath = `M ${-needleBaseWidth/2} 0 
                       L 0 ${-needleLength} 
                       L ${needleBaseWidth/2} 0 
                       Z`;

  chartGroup.append('path')
    .attr('d', needlePath)
    .attr('transform', `rotate(${(needleAngle * 180 / Math.PI) + 90})`)
    .attr('fill', '#000')
    .style('filter', 'drop-shadow(2px 2px 4px rgba(0,0,0,0.4))');

  // Needle base (center circle)
  chartGroup.append('circle')
    .attr('r', 15)
    .attr('fill', '#000');
    
  chartGroup.append('circle')
    .attr('r', 10)
    .attr('fill', '#333');

  // 12. CENTER VALUE DISPLAY
  chartGroup.append('text')
    .attr('text-anchor', 'middle')
    .attr('y', 10)
    .style('font-size', '46px')
    .style('font-weight', 'bold')
    .style('fill', palette[0])
    .text('$' + d3.format(',')(Math.round(totalValue)));

  // Optional: Add percentage or measure name below value
  const percentOfMax = ((totalValue / maxScale) * 100).toFixed(1);
  chartGroup.append('text')
    .attr('text-anchor', 'middle')
    .attr('y', 50)
    .style('font-size', '14px')
    .style('fill', '#666')
    .text(`${percentOfMax}% of max`);

  // 13. DEBUG INFO (console)
  console.log(`Needle pointing to: ${totalValue} (${percentOfMax}% of ${maxScale})`);
  console.log(`Blue arc ends at angle: ${(valueEndAngle * 180 / Math.PI).toFixed(1)}°`);

  // 14. INTERACTION
  const interactionRect = chartGroup.append('rect')
    .attr('x', -radius)
    .attr('y', -radius)
    .attr('width', radius * 2)
    .attr('height', radius * 2)
    .attr('fill', 'transparent')
    .style('cursor', 'pointer');

  interactionRect.datum({ 
      allTupleIds: allTupleIds,
      value: totalValue 
  });

  return { 
      viz: svg.node(), 
      interactionElement: interactionRect, 
      allTupleIds: allTupleIds
  };
}

// HELPER FUNCTIONS
async function renderViz (rawData, encodingMap, selectedMarksIds, styles) {
  const encodedData = getEncodedData(rawData, encodingMap);
  const content = document.getElementById('content');
  content.innerHTML = '';
  const gaugeChart = await GaugeChart(encodedData, encodingMap, content.offsetWidth, content.offsetHeight, selectedMarksIds, styles);
  if (gaugeChart.viz) content.appendChild(gaugeChart.viz);
  return gaugeChart;
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
    console.log('✓ Gauge Extension Initialized');
    
    const worksheet = tableau.extensions.worksheetContent.worksheet;
    let summaryData = {}, encodingMap = {}, selectedMarks = new Map(), hoveredMarks = new Map();
    let interactionElement = null;
    let allTupleIds = [];
    const styles = tableau.extensions.environment.workbookFormatting?.formattingSheets?.find(x => x.classNameKey === 'tableau-worksheet')?.cssProperties;

    const updateDataAndRender = async () => {
      try {
        [summaryData, encodingMap] = await Promise.all([getSummaryDataTable(worksheet), getEncodingMap(worksheet)]);
        selectedMarks = await getSelection(worksheet, summaryData);
        const result = await renderViz(summaryData, encodingMap, selectedMarks, styles);
        interactionElement = result.interactionElement;
        allTupleIds = result.allTupleIds;
      } catch (error) {
        console.error(error);
        document.getElementById('content').innerHTML = '<div style="padding:40px;color:red;">Error: ' + error.message + '</div>';
      }
    };

    onresize = async () => {
      if (summaryData && Object.keys(summaryData).length > 0) {
        const result = await renderViz(summaryData, encodingMap, selectedMarks, styles);
        interactionElement = result.interactionElement;
        allTupleIds = result.allTupleIds;
      }
    };
    
    worksheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, async () => {
      console.log('📊 Data changed - updating gauge...');
      await updateDataAndRender();
    });
    
    worksheet.addEventListener(tableau.TableauEventType.FilterChanged, async () => {
      console.log('🔍 Filter changed - updating gauge...');
      await updateDataAndRender();
    });
    
    document.body.addEventListener('click', async (e) => {
      const elem = d3.select(document.elementFromPoint(e.pageX, e.pageY));
      const data = elem?.datum();
      if (data && data.allTupleIds) {
          if (!e.ctrlKey) selectedMarks.clear();
          const allSelected = data.allTupleIds.every(id => selectedMarks.has(id));
          if (allSelected) selectedMarks.clear();
          else data.allTupleIds.forEach(id => selectedMarks.set(id, true));
          selectTuples(e.pageX, e.pageY, selectedMarks, hoveredMarks);
          const result = await renderViz(summaryData, encodingMap, selectedMarks, styles);
          interactionElement = result.interactionElement;
          allTupleIds = result.allTupleIds;
      } else if (!e.ctrlKey && selectedMarks.size > 0) {
          selectedMarks.clear();
          selectTuples(e.pageX, e.pageY, selectedMarks, hoveredMarks);
          const result = await renderViz(summaryData, encodingMap, selectedMarks, styles);
          interactionElement = result.interactionElement;
          allTupleIds = result.allTupleIds;
      }
    });
    
    document.body.addEventListener('mousemove', e => {
        const elem = d3.select(document.elementFromPoint(e.pageX, e.pageY));
        const data = elem?.datum();
        if (data && data.allTupleIds) {
            const representativeId = data.allTupleIds[0];
            if (!hoveredMarks.has(representativeId)) {
                hoveredMarks.clear();
                hoveredMarks.set(representativeId);
                getWorksheet().hoverTupleAsync(representativeId, { tooltipAnchorPoint: { x: e.pageX, y: e.pageY } });
            }
        } else if (hoveredMarks.size > 0) {
            clearHoveredMarks(hoveredMarks);
            getWorksheet().hoverTupleAsync(null, { tooltipAnchorPoint: { x: e.pageX, y: e.pageY } });
        }
    });

    document.body.addEventListener('mouseout', () => clearHoveredMarks(hoveredMarks));
    await updateDataAndRender();
  }).catch (err => {
    console.error(err);
    document.getElementById('content').innerHTML = '<div style="padding:40px;color:red;">Error: ' + err.message + '</div>';
  });
};

function selectTuples (x, y, selectedTupleIds, hoveredTupleIds) {
  clearHoveredMarks(hoveredTupleIds);
  getWorksheet().selectTuplesAsync([...selectedTupleIds.keys()], tableau.SelectOptions.Simple, { tooltipAnchorPoint: { x, y } });
}
function clearHoveredMarks (hoveredTupleIds) { hoveredTupleIds.clear(); }
async function getSelection (worksheet, allMarks) {
  try { return findIdsOfSelectedMarks(allMarks, await worksheet.getSelectedMarksAsync()); } catch (e) { return new Map(); }
}
function findIdsOfSelectedMarks (allMarks, selectedMarks) {
  if (!selectedMarks || !selectedMarks.data || !selectedMarks.data.length) return new Map();
  const columns = selectedMarks.data[0].columns;
  const selectedMarkMap = new Map();
  const selectedMarksIds = new Map();
  for (const sm of convertToListOfNamedRows(selectedMarks.data[0])) {
    let key = ''; for (const col of columns) key += sm[col.fieldName].value + '\x00';
    selectedMarkMap.set(key, sm);
  }
  let tupleId = 1;
  for (const mark of allMarks) {
    let key = ''; for (const col of columns) key += mark[col.fieldName].value + '\x00';
    if (selectedMarkMap.has(key)) selectedMarksIds.set(tupleId);
    tupleId++;
  }
  return selectedMarksIds;
}
function convertToListOfNamedRows (dataTablePage) {
  const rows = []; const columns = dataTablePage.columns; const data = dataTablePage.data;
  for (let i = 0; i < data.length; ++i) {
    const row = {};
    for (let j = 0; j < columns.length; ++j) row[columns[j].fieldName] = data[i][columns[j].index];
    row.tupleId = i + 1; rows.push(row);
  }
  return rows;
}
async function getSummaryDataTable (worksheet) {
  let rows = [];
  const dataTableReader = await worksheet.getSummaryDataReaderAsync(undefined, { ignoreSelection: true });
  for (let i = 0; i < dataTableReader.pageCount; i++) rows = rows.concat(convertToListOfNamedRows(await dataTableReader.getPageAsync(i)));
  await dataTableReader.releaseAsync();
  return rows;
}
function getEncodedData (data, encodingMap) {
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
function getWorksheet () {
  return tableau.extensions.worksheetContent ? tableau.extensions.worksheetContent.worksheet : tableau.extensions.dashboardContent.dashboard.worksheets[0];
}
/* global d3 */
/* global tinycolor */
/* global tableau */

const backgroundColor = tinycolor('white');
const palette = ['#5B6FD8', '#D3D3D3', '#4e79a7', '#f28e2c'];

// 👇 YAHAN SE APNI SETTINGS BADALEN 👇
// Default Config
const DEFAULT_CONFIG = {
    measure: "Sales", 
    format: "#,##0.00", 
    useSuffix: true, 
    decimalPlaces: 1, 
    prefix: "", 
    unit: " KWh" 
};

// Load Settings from LocalStorage or use Default
let CONFIG = JSON.parse(localStorage.getItem('gaugeConfig')) || DEFAULT_CONFIG;
// 👆 YAHAN SE SETTINGS KHATAM 👆

// Purani line hata di gayi hai: const CUSTOM_UNIT = " KW";

function formatNumber(value) {
    if (!CONFIG.useSuffix) {
        return value.toFixed(CONFIG.decimalPlaces);
    }
    if (value >= 1000000) {
        return (value / 1000000).toFixed(CONFIG.decimalPlaces) + "M";
    } else if (value >= 1000) {
        return (value / 1000).toFixed(CONFIG.decimalPlaces) + "K";
    } else {
        return Math.round(value);
    }
}

// MAIN GAUGE CHART FUNCTION
async function GaugeChart(encodedData, encodingMap, width, height, selectedTupleIds, styles) {
  
  let valueKey = null;
  let targetKey = null;
  let totalValue, totalTarget, maxScale, allTupleIds;

  if (window._lockedFinalValues) {
    totalValue = window._lockedFinalValues.totalValue;
    totalTarget = window._lockedFinalValues.totalTarget;
    maxScale = window._lockedFinalValues.maxScale;
    allTupleIds = window._lockedFinalValues.allTupleIds;
    if (totalTarget > 0) targetKey = 'dummy'; 
  } else {
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
          if (numericKeys.includes(CONFIG.measure)) {
              valueKey = CONFIG.measure;
              targetKey = numericKeys.find(k => k !== CONFIG.measure);
          } else if (foundTarget && numericKeys.length > 1) {
              targetKey = foundTarget;
              valueKey = numericKeys.find(k => k !== foundTarget);
          } else if (numericKeys.length > 1) {
    valueKey = numericKeys[0];
    targetKey = numericKeys[1];
} else {
    valueKey = numericKeys[0];
    targetKey = null;
}
      }

      if (!valueKey) return { viz: null };

      totalValue = 0;
      totalTarget = 0;
      allTupleIds = [];
      
      encodedData.forEach(row => {
        totalValue += parseFloat(row[valueKey]?.[0]?.value || 0);
        if (targetKey) {
            totalTarget = parseFloat(row[targetKey]?.[0]?.value || 0);
        }
        if (row.tupleId) allTupleIds.push(row.tupleId);
      });

      if (totalTarget > 0) {
          const rawMax = totalTarget * 1.25;
          const pow10 = Math.pow(10, Math.floor(Math.log10(rawMax)));
          maxScale = Math.ceil(rawMax / (pow10 / 2)) * (pow10 / 2);
      } else {
          const rawMax = totalValue * 1.4;
          const pow10 = Math.pow(10, Math.floor(Math.log10(rawMax)));
          maxScale = Math.ceil(rawMax / (pow10 / 2)) * (pow10 / 2);
      }

      window._lockedFinalValues = {
          totalValue: totalValue,
          totalTarget: totalTarget,
          maxScale: maxScale,
          allTupleIds: allTupleIds
      };
  }

  width = Math.max(width, 100);
  height = Math.max(height, 100);

  const minDim = Math.min(width, height);
  const margin = { top: minDim * 0.12, right: minDim * 0.15, bottom: minDim * 0.12, left: minDim * 0.15 };
  
  const cx = width / 2;
  const cy = height / 2;
  const radius = Math.min((width - margin.left - margin.right) / 2.6, (height - margin.top - margin.bottom) / 2.15);

  const svg = d3.create('svg')
    .attr('class', tableau.ClassNameKey.Worksheet)
    .attr('width', '100%')
    .attr('height', '100%')
    .attr('viewBox', `0 0 ${width} ${height}`)
    .attr('preserveAspectRatio', 'xMidYMid meet')
    .style('background', 'white')
    .style('display', 'block');
    
  const chartGroup = svg.append('g')
    .attr('transform', `translate(${cx}, ${cy})`);

  // Angles
  const startAngle = -Math.PI * 0.75; 
  const endAngle = Math.PI * 0.75;   
  const totalRange = endAngle - startAngle;
  const valueFraction = Math.min(Math.max(totalValue / maxScale, 0), 1);
  const currentAngle = startAngle + (valueFraction * totalRange);

  // Arcs
  const arcGen = d3.arc()
    .innerRadius(radius * 0.7)
    .outerRadius(radius)
    .cornerRadius(Math.max(2, radius * 0.05));
    
  chartGroup.append('path')
    .attr('d', arcGen({ startAngle, endAngle }))
    .attr('fill', '#D3D3D3');
    
  chartGroup.append('path')
.attr('d', arcGen({ startAngle, endAngle: currentAngle + 0.03 }))
    .attr('fill', palette[0]);

  // Ticks
  const numLabels = 5;
  for (let i = 0; i <= numLabels; i++) {
      const f = i / numLabels;
      const angle = startAngle + (f * totalRange);
      const x = Math.sin(angle); 
      const y = -Math.cos(angle);
      const tickStart = radius * 1.05;
      const tickEnd = radius * 1.15;
      const labelR = radius * 1.32;

      chartGroup.append('line')
        .attr('x1', tickStart * x)
        .attr('y1', tickStart * y)
        .attr('x2', tickEnd * x)
        .attr('y2', tickEnd * y)
        .attr('stroke', '#333')
        .attr('stroke-width', Math.max(1, radius * 0.02));
      
      chartGroup.append('text')
        .attr('x', labelR * x)
        .attr('y', labelR * y)
        .attr('text-anchor', 'middle')
        .attr('dominant-baseline', 'middle')
        .style('font-family', 'Arial, sans-serif')
        .style('font-size', Math.max(9, radius * 0.15) + 'px')
        .style('fill', '#333')
        .style('font-weight', 'bold')
       .text(formatNumber(f * maxScale) + CONFIG.unit);
  }

  // Target line
  if (targetKey && totalTarget > 0) {
      const tAngle = startAngle + (Math.min(totalTarget / maxScale, 1) * totalRange);
      chartGroup.append('line')
        .attr('x1', radius * 0.6 * Math.sin(tAngle))
        .attr('y1', radius * 0.6 * -Math.cos(tAngle))
        .attr('x2', radius * 1.15 * Math.sin(tAngle))
        .attr('y2', radius * 1.15 * -Math.cos(tAngle))
        .attr('stroke', '#f28e2c')
        .attr('stroke-width', Math.max(2, radius * 0.04))
        .attr('stroke-linecap', 'round');
  }

  // Needle
  const needleGroup = chartGroup.append('g')
    .attr('transform', `rotate(${(currentAngle * 180 / Math.PI)})`);
  const needleLen = radius * 0.85;
  const needleWidth = Math.max(4, radius * 0.04);
  
  needleGroup.append('path')
    .attr('d', `M -${needleWidth} 0 L 0 -${needleLen} L ${needleWidth} 0 Z`)
    .attr('fill', '#333');
    
  chartGroup.append('circle')
    .attr('r', Math.max(3, radius * 0.06))
    .attr('fill', '#333');
  
  chartGroup.append('circle')
    .attr('r', Math.max(1.5, radius * 0.03))
    .attr('fill', '#fff');

  // ANIMATED VALUE
   const valueText = chartGroup.append('text')
    .attr('y', radius * 0.30)
    .attr('text-anchor', 'middle')
    .attr('dominant-baseline', 'middle')
    .style('font-family', 'Arial, sans-serif')
    .style('font-size', Math.max(16, radius * 0.25) + 'px')
    .style('font-weight', 'bold')
    .style('fill', palette[0])
    .text(formatNumber(0) + CONFIG.unit);
valueText.raise();

  
  const animationDuration = 1000;
  const startTime = Date.now();
  
  function animateValue() {
    const elapsed = Date.now() - startTime;
    const progress = Math.min(elapsed / animationDuration, 1);
    const easeProgress = 1 - Math.pow(1 - progress, 3);
    const currentValue = totalValue * easeProgress;
    
    valueText.text(CONFIG.prefix + formatNumber(currentValue) + CONFIG.unit);

    if (progress < 1) requestAnimationFrame(animateValue);
  }
  requestAnimationFrame(animateValue);

  // Percentage
  let actualPercentage = 0;
  if (totalTarget > 0) {
      actualPercentage = (totalValue / totalTarget) * 100;
  } else {
      actualPercentage = totalValue > 0 ? 100 : 0; 
  }

  const percentageText = chartGroup.append('text')
    .attr('y', radius * 0.70)
    .attr('text-anchor', 'middle')
    .attr('dominant-baseline', 'middle')
    .style('font-family', 'Arial, sans-serif')
    .style('font-size', Math.max(10, radius * 0.15) + 'px')
    .style('font-weight', 'bold')
    .style('fill', '#666')
    .text(`${actualPercentage.toFixed(1)}% achieved`);

percentageText.raise();

  // Target
if (targetKey && totalTarget > 0) {
    const targetText = chartGroup.append('text')
      .attr('y', radius * 0.85)
      .attr('text-anchor', 'middle')
      .attr('dominant-baseline', 'middle')
      .style('font-family', 'Arial, sans-serif')
      .style('font-size', Math.max(8, radius * 0.12) + 'px')
            .style('font-weight', '400')
      .style('fill', '#999')
      .text('Target: ' + formatNumber(totalTarget) + CONFIG.unit);

    targetText.raise();
}

  // 👉 SETTINGS ICON (SVG)
  const settingsIcon = chartGroup.append('g')
    .attr('transform', `translate(${radius * 0.8}, ${-radius * 0.8})`)
    .style('cursor', 'pointer');
  
  settingsIcon.append('circle')
    .attr('r', 15)
    .attr('fill', '#f0f0f0')
    .attr('stroke', '#ccc')
    .attr('stroke-width', 1);
  
  settingsIcon.append('path')
    .attr('d', 'M12 15.5A3.5 3.5 0 0 1 8.5 12 3.5 3.5 0 0 1 12 8.5a3.5 3.5 0 0 1 3.5 3.5 3.5 3.5 0 0 1-3.5 3.5m7.43-2.53c.04-.32.07-.64.07-.97 0-.33-.03-.66-.07-1l2.11-1.63c.19-.15.24-.42.12-.64l-2-3.46c-.12-.22-.39-.31-.61-.22l-2.49 1c-.52-.39-1.06-.73-1.69-.98l-.37-2.65A.506.506 0 0 0 14 2h-4c-.25 0-.46.18-.5.42l-.37 2.65c-.63.25-1.17.59-1.69.98l-2.49-1c-.22-.09-.49 0-.61.22l-2 3.46c-.13.22-.07.49.12.64L4.57 11c-.04.33-.07.65-.07.97 0 .32.03.65.07.97l-2.11 1.66c-.19.15-.25.42-.12.64l2 3.46c.12.22.39.3.61.22l2.49-1.01c.52.4 1.06.74 1.69.99l.37 2.65c.04.24.25.42.5.42h4c.25 0 .46-.18.5-.42l.37-2.65c.63-.24 1.17-.59 1.69-.99l2.49 1.01c.22.08.49 0 .61-.22l2-3.46c.12-.22.07-.49-.12-.64l-2.11-1.66z')
    .attr('fill', '#555')
    .attr('transform', 'scale(0.8) translate(6,6)');

  // 👉 SETTINGS MODAL (HTML/CSS)
  const modal = document.createElement('div');
  modal.id = 'gaugeSettingsModal';
  modal.style.cssText = `
    position: absolute;
    top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(0,0,0,0.5);
    display: none;
    justify-content: center;
    align-items: center;
    z-index: 1000;
  `;
  
  const modalContent = document.createElement('div');
  modalContent.style.cssText = `
    background: white;
    padding: 20px;
    border-radius: 8px;
    width: 300px;
    font-family: Arial, sans-serif;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
  `;
  
  modalContent.innerHTML = `
    <h3 style="margin-top:0; border-bottom: 1px solid #eee; padding-bottom:10px;">& Gauge Settings</h3>
    
    <label style="display:block; margin-bottom:5px; font-weight:bold;">Decimal Places</label>
    <input type="number" id="setDecimal" value="${CONFIG.decimalPlaces}" min="0" max="5" style="width:100%; padding:5px; margin-bottom:15px; box-sizing:border-box;">
    
    <label style="display:block; margin-bottom:5px; font-weight:bold;">Unit Label</label>
    <input type="text" id="setUnit" value="${CONFIG.unit}" style="width:100%; padding:5px; margin-bottom:15px; box-sizing:border-box;">
    
    <label style="display:block; margin-bottom:5px; font-weight:bold;">Prefix (e.g., ₹, $)</label>
    <input type="text" id="setPrefix" value="${CONFIG.prefix}" style="width:100%; padding:5px; margin-bottom:15px; box-sizing:border-box;">
    
    <label style="display:flex; align-items:center; justify-content:space-between;">
      <span style="font-weight:bold;">Use Suffix (K/M)</span>
      <input type="checkbox" id="setSuffix" ${CONFIG.useSuffix ? 'checked' : ''} style="transform:scale(1.2);">
    </label>
    
    <div style="margin-top:20px; text-align:right;">
      <button onclick="resetSettings()" style="padding:5px 10px; cursor:pointer; background:#eee; border:1px solid #ccc;">Reset</button>
      <button onclick="saveSettings()" style="padding:5px 10px; cursor:pointer; background:#007bff; color:white; border:none; margin-left:5px;">Save</button>
    </div>
  `;
  
  modal.appendChild(modalContent);
  svg.node().parentNode.appendChild(modal);

  // 👉 EVENT LISTENERS
  settingsIcon.on('click', () => {
    document.getElementById('gaugeSettingsModal').style.display = 'flex';
  });
  
  modal.on('click', (e) => {
    if (e.target === modal) {
      modal.style.display = 'none';
    }
  });

  // 👉 GLOBAL FUNCTIONS FOR MODAL
  window.saveSettings = function() {
    CONFIG.decimalPlaces = parseInt(document.getElementById('setDecimal').value);
    CONFIG.unit = document.getElementById('setUnit').value;
    CONFIG.prefix = document.getElementById('setPrefix').value;
    CONFIG.useSuffix = document.getElementById('setSuffix').checked;
    
    localStorage.setItem('gaugeConfig', JSON.stringify(CONFIG));
    
    // Redraw chart
    renderViz(encodedData, encodingMap, selectedTupleIds, styles);
    
    document.getElementById('gaugeSettingsModal').style.display = 'none';
  };
  
  window.resetSettings = function() {
    CONFIG = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
    localStorage.removeItem('gaugeConfig');
    
    // Reload page to reset
    location.reload();
  };

  // Interaction
  const interactionElement = chartGroup.append('circle')
    .attr('r', radius)
    .attr('fill', 'transparent')
    .style('cursor', 'pointer');
    
  interactionElement.datum({ allTupleIds, value: totalValue });
  
  return { viz: svg.node(), interactionElement, allTupleIds };
}

// RENDER
async function renderViz(rawData, encodingMap, selectedMarksIds, styles) {
  const encodedData = getEncodedData(rawData, encodingMap);
  const content = document.getElementById('content');
  
  if (!content) {
    console.error('Content div not found!');
    return { viz: null };
  }
  
  content.innerHTML = '';
  
  const width = content.offsetWidth || content.clientWidth || 400;
  const height = content.offsetHeight || content.clientHeight || 300;
  
  const result = await GaugeChart(encodedData, encodingMap, width, height, selectedMarksIds, styles);
  
  if (result.viz) content.appendChild(result.viz);
  return result;
}

// INIT
window.onload = function() {
  tableau.extensions.initializeAsync().then(async () => {
    window._lockedFinalValues = null;
    const worksheet = getWorksheet();
    let summaryData = {};
    let encodingMap = {};
    let selectedMarks = new Map();
    const styles = tableau.extensions.environment.workbookFormatting?.formattingSheets?.find(
      x => x.classNameKey === 'tableau-worksheet'
    )?.cssProperties;

    const update = async () => {
      window._lockedFinalValues = null;
      [summaryData, encodingMap] = await Promise.all([
        getSummaryDataTable(worksheet), 
        getEncodingMap()
      ]);
      selectedMarks = await getSelection(worksheet, summaryData);
      await renderViz(summaryData, encodingMap, selectedMarks, styles);
    };

    worksheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, update);
    worksheet.addEventListener(tableau.TableauEventType.FilterChanged, update);

    let resizeTimeout;
    const resizeObserver = new ResizeObserver(() => {
      clearTimeout(resizeTimeout);
      resizeTimeout = setTimeout(() => {
        const content = document.getElementById('content');
        if (content && summaryData && encodingMap) {
          renderViz(summaryData, encodingMap, selectedMarks, styles);
        }
      }, 150);
    });
    
    const contentDiv = document.getElementById('content');
    if (contentDiv) resizeObserver.observe(contentDiv);

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

// HELPERS
async function getEncodingMap() {
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
  for (let i = 0; i < reader.pageCount; i++) {
    rows = rows.concat(convertToListOfNamedRows(await reader.getPageAsync(i)));
  }
  await reader.releaseAsync();
  return rows;
}

function convertToListOfNamedRows(dataTablePage) {
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

function getEncodedData(data, encodingMap) {
  const encodedData = [];
  let tupleId = 1;
  for (const row of data) {
    const encodedRow = {};
    for (const encName in encodingMap) {
      encodedRow[encName] = [];
      for (const field of encodingMap[encName]) {
        encodedRow[encName].push(row[field.name]);
      }
    }
    encodedRow.tupleId = tupleId++;
    encodedData.push(encodedRow);
  }
  return encodedData;
}

function getWorksheet() {
  return tableau.extensions.worksheetContent 
    ? tableau.extensions.worksheetContent.worksheet 
    : tableau.extensions.dashboardContent.dashboard.worksheets[0];
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
        if (columns.map(col => m[col.fieldName].value).join('|') === key) {
          selectedMarksIds.set(idx + 1);
        }
      });
    }
    return selectedMarksIds;
  } catch (e) { 
    return new Map(); 
  }
}
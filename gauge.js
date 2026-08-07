/* global d3 */
/* global tinycolor */
/* global tableau */

const backgroundColor = tinycolor('white');
const palette = ['#5B6FD8', '#D3D3D3', '#4e79a7', '#f28e2c'];

let config = {};
try {
    if (styles && styles.configJson) {
        config = JSON.parse(styles.configJson);
    }
} catch (e) {
    console.error("Error parsing config:", e);
}

const CONFIG = {
    measure: config.measure || "Sales",
    format: config.format || "#,##0.00",
    useSuffix: config.useSuffix !== false,
    decimalPlaces: config.decimalPlaces || 2,
    decimalPlacesForAchievedValue: config.decimalPlacesForAchievedValue || 2,
    prefix: config.prefix || "",
    unit: "",                               // No unit (KWh removed completely)
    isPercentage: config.isPercentage || false
};

// Global mode state initialization
window.currentGaugeMode = window.currentGaugeMode || "value";

function formatNumber(value, useSuffix = CONFIG.useSuffix, isPercentage = CONFIG.isPercentage, includeUnit = true, decimals = null) {
    value = Number(value) || 0;
    const dMain = (decimals != null) ? decimals : CONFIG.decimalPlaces;
    const dVal  = (decimals != null) ? decimals : CONFIG.decimalPlacesForAchievedValue;

    if (isPercentage) {
        const formattedPct = value.toLocaleString('en-US', {
            minimumFractionDigits: dMain,
            maximumFractionDigits: dMain
        });
        return formattedPct + "%";
    }

    if (!useSuffix) {
        return value.toLocaleString('en-US', {
            minimumFractionDigits: dMain,
            maximumFractionDigits: dMain
        });
    }

    let formattedValue;
    if (Math.abs(value) >= 1000000) {
        formattedValue = (value / 1000000).toFixed(dVal) + "M";
    } else if (Math.abs(value) >= 1000) {
        formattedValue = (value / 1000).toFixed(dVal) + "K";
    } else {
        formattedValue = value.toFixed(dVal);
    }

    return (CONFIG.prefix || "") + formattedValue + (includeUnit ? (CONFIG.unit || "") : "");
}

// ✨ AUTO-FIT: calculates precise font size dynamically
function fitFontSize(text, maxWidth, baseSize, minSize = 8) {
    const len = String(text).length || 1;
    const approx = maxWidth / (len * 0.56);  
    return Math.max(minSize, Math.min(baseSize, approx));
}


async function GaugeChart(encodedData, encodingMap, width, height, selectedMarksIds, styles) {
    let valueKey = null;
    let targetKey = null;
    let totalValue, totalTarget, maxScale, allTupleIds;

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
            totalTarget += parseFloat(row[targetKey]?.[0]?.value || 0);
        }
        if (row.tupleId) allTupleIds.push(row.tupleId);
    });

    const isPctMode = (window.currentGaugeMode === "percentage");
    const achievedPct = totalTarget > 0 ? (totalValue / totalTarget) * 100 : 0;

    // 1️⃣ SCALE CALCULATOR: प्रतिशत मोड होने पर पैमाना हमेशा टारगेट से 20% आगे (120%) रहेगा ताकि टारगेट हमेशा साफ़ दिखे
    if (isPctMode) {
        maxScale = achievedPct > 100 ? Math.ceil((achievedPct * 1.2) / 10) * 10 : 120; 
    } else {
        const rawMax = (Math.max(totalValue, totalTarget) * 1.2) || 100;
        const pow10 = Math.pow(10, Math.floor(Math.log10(rawMax)));
        maxScale = Math.ceil(rawMax / (pow10 / 2)) * (pow10 / 2);
    }

    const sampleFormatted = isPctMode ? "120.00%" : formatNumber(maxScale, false, false, true, CONFIG.decimalPlaces);
    const numberLength = sampleFormatted.length; 
    const fontScale = Math.min(1.0, Math.max(0.7, 10 / numberLength));

    width = Math.max(width, 100);
    height = Math.max(height, 100);

    const minDim = Math.min(width, height);
    const margin = {
        top: minDim * 0.12,
        right: numberLength > 10 ? minDim * 0.20 : minDim * 0.15,
        bottom: minDim * 0.12,
        left: numberLength > 10 ? minDim * 0.20 : minDim * 0.15
    };

    const cx = width / 2;
    const cy = height / 2;
    
    const radiusDivisor = numberLength > 10 ? 3.0 : (numberLength > 7 ? 2.8 : 2.6);
    const radius = Math.min(
        (width - margin.left - margin.right) / radiusDivisor,
        (height - margin.top - margin.bottom) / 2.15
    );

    const svg = d3.create('svg')
        .attr('class', tableau.ClassNameKey.Worksheet)
        .attr('width', '100%')
        .attr('height', '100%')
        .attr('viewBox', `0 0 ${width} ${height}`)
        .attr('preserveAspectRatio', 'xMidYMid meet')
        .style('background', 'white');

    const chartGroup = svg.append('g')
        .attr('transform', `translate(${cx}, ${cy})`);

    const startAngle = -Math.PI * 0.75;
    const endAngle = Math.PI * 0.75;
    const totalRange = endAngle - startAngle;

    // 2️⃣ NEEDLE POSITION: सुई हमेशा अचीव्ड% के हिसाब से घूमेगी
    let valueFraction;
    if (isPctMode) {
        valueFraction = Math.min(Math.max(achievedPct / maxScale, 0), 1);
    } else {
        valueFraction = Math.min(Math.max(totalValue / maxScale, 0), 1);
    }
    const currentAngle = startAngle + (valueFraction * totalRange);

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

    const numLabels = 1;
    for (let i = 0; i <= numLabels; i++) {
        const f = i / numLabels;
        const angle = startAngle + (f * totalRange);
        const x = Math.sin(angle);
        const y = -Math.cos(angle);
        const tickStart = radius * 1.05;
        const tickEnd = radius * 1.15;
        const labelR = radius * 1.22; 

        chartGroup.append('line')
            .attr('x1', tickStart * x)
            .attr('y1', tickStart * y)
            .attr('x2', tickEnd * x)
            .attr('y2', tickEnd * y)
            .attr('stroke', '#333')
            .attr('stroke-width', Math.max(1, radius * 0.02));

        const tickStr = formatNumber(f * maxScale, false, isPctMode, true, CONFIG.decimalPlaces);
        const isLeft = (i === 0);
        const textAnchor = isLeft ? 'end' : 'start';
        const xOffset = isLeft ? -5 : 5;

        chartGroup.append('text')
            .attr('x', labelR * x + xOffset)
            .attr('y', labelR * y)
            .attr('text-anchor', textAnchor)
            .attr('dominant-baseline', 'middle')
            .style('font-family', 'Arial, sans-serif')
            .style('font-size', fitFontSize(tickStr, radius * 0.85, radius * 0.13) * fontScale + 'px')
            .style('fill', '#333')
            .style('font-weight', 'bold')
            .text(tickStr);
    }

    // 3️⃣ TARGET ORANGE LINE: प्रतिशत मोड में टारगेट लाइन हमेशा गेज के अंदर '100%' पर सुंदर दिखेगी
    if (targetKey && totalTarget > 0) {
        const targetPosFraction = isPctMode ? (100 / maxScale) : (totalTarget / maxScale);
        const tAngle = startAngle + (Math.min(targetPosFraction, 1) * totalRange);
        chartGroup.append('line')
            .attr('x1', radius * 0.6 * Math.sin(tAngle))
            .attr('y1', radius * 0.6 * -Math.cos(tAngle))
            .attr('x2', radius * 1.15 * Math.sin(tAngle))
            .attr('y2', radius * 1.15 * -Math.cos(tAngle))
            .attr('stroke', '#f28e2c')
            .attr('stroke-width', Math.max(2, radius * 0.04))
            .attr('stroke-linecap', 'round');
    }

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

    // 4️⃣ CENTER VALUE: यह हमेशा बिना K/M के पूरा वास्तविक नंबर ही रहेगा (दोनों मोड में)
    const centerStr = formatNumber(totalValue, false, false, false, CONFIG.decimalPlacesForAchievedValue);
    const valueText = chartGroup.append('text')
        .attr('y', radius * 0.22) 
        .attr('text-anchor', 'middle')
        .attr('dominant-baseline', 'middle')
        .style('font-family', 'Arial, sans-serif')
        .style('font-size', fitFontSize(centerStr, radius * 1.25, radius * 0.22) * fontScale + 'px')
        .style('font-weight', 'bold')
        .style('fill', palette[0])
        .text(centerStr);
    valueText.raise();

    const animationDuration = 1000;
    const startTime = Date.now();

    function animateValue() {
        const elapsed = Date.now() - startTime;
        const progress = Math.min(elapsed / animationDuration, 1);
        const easeProgress = 1 - Math.pow(1 - progress, 3);
        const currentValue = totalValue * easeProgress;
        valueText.text(formatNumber(currentValue, false, false, false, CONFIG.decimalPlacesForAchievedValue));
        if (progress < 1) requestAnimationFrame(animateValue);
    }
    requestAnimationFrame(animateValue);

    // % ACHIEVED
    const pctStr = `${achievedPct.toFixed(2)}% achieved`;
    const percentageText = chartGroup.append('text')
        .attr('y', radius * 0.48) 
        .attr('text-anchor', 'middle')
        .attr('dominant-baseline', 'middle')
        .style('font-family', 'Arial, sans-serif')
        .style('font-size', fitFontSize(pctStr, radius * 1.4, radius * 0.13) * fontScale + 'px')
        .style('font-weight', 'bold')
        .style('fill', '#666')
        .text(pctStr);
    percentageText.raise();

    // 5️⃣ TARGET BOTTOM TEXT: प्रतिशत मोड होने पर नीचे टारगेट 'Target: 100.00%' दिखाएगा
    if (targetKey && totalTarget > 0) {
        let targetStr;
        if (isPctMode) {
            targetStr = 'Target: ' + formatNumber(100, false, true, true, CONFIG.decimalPlaces);
        } else {
            targetStr = 'Target: ' + formatNumber(totalTarget, false, false, true, CONFIG.decimalPlaces);
        }

        const targetText = chartGroup.append('text')
            .attr('y', radius * 0.72) 
            .attr('text-anchor', 'middle')
            .attr('dominant-baseline', 'middle')
            .style('font-family', 'Arial, sans-serif')
            .style('font-size', fitFontSize(targetStr, radius * 1.5, radius * 0.11) * fontScale + 'px')
            .style('font-weight', '400')
            .style('fill', '#999')
            .text(targetStr);
        targetText.raise();
    }

    const interactionElement = chartGroup.append('circle')
        .attr('r', radius)
        .attr('fill', 'transparent')
        .style('cursor', 'pointer');

    interactionElement.datum({ allTupleIds, value: totalValue });

    return { viz: svg.node(), interactionElement, allTupleIds };
}

// ... बाकी का सारा Tableau रेंडरिंग कोड समान रहेगा ...
async function renderViz(rawData, encodingMap, selectedMarksIds, styles) {
    const encodedData = getEncodedData(rawData, encodingMap);
    const content = document.getElementById('content');
    if (!content) {
        console.error('Content div not found!');
        return { viz: null };
    }

    window.gaugeActiveArgs = { rawData, encodingMap, selectedMarksIds, styles };

    // BUTTON CONTROLS BUILDER
    if (!document.getElementById('gauge-controls-container')) {
        content.innerHTML = `
            <div id="gauge-controls-container" style="display: flex; gap: 16px; align-items: center; justify-content: center; padding: 10px; background: #ffffff; border-bottom: 1px solid #e2e8f0; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Arial, sans-serif; font-size: 13px; box-shadow: 0 1px 3px rgba(0,0,0,0.04);">
                <span style="font-weight: 600; color: #2d3748; letter-spacing: 0.3px;">View Format:</span>
                <div style="display: flex; gap: 14px; align-items: center;">
                    <label style="cursor: pointer; color: #4a5568; display: flex; align-items: center; gap: 6px; font-weight: 500; transition: color 0.2s;">
                        <input type="radio" name="gaugeMode" value="value" ${window.currentGaugeMode !== 'percentage' ? 'checked' : ''} style="accent-color: #5B6FD8; width: 15px; height: 15px; cursor: pointer;"> Raw Number
                    </label>
                    <label style="cursor: pointer; color: #4a5568; display: flex; align-items: center; gap: 6px; font-weight: 500; transition: color 0.2s;">
                        <input type="radio" name="gaugeMode" value="percentage" ${window.currentGaugeMode === 'percentage' ? 'checked' : ''} style="accent-color: #5B6FD8; width: 15px; height: 15px; cursor: pointer;"> Percentage (%)
                    </label>
                </div>
                <button id="apply-gauge-btn" style="background: #5B6FD8; color: white; border: none; padding: 6px 16px; border-radius: 6px; cursor: pointer; font-weight: 600; font-size: 12px; transition: all 0.2s; box-shadow: 0 1px 3px rgba(91, 111, 216, 0.3); outline: none;">Apply</button>
            </div>
            <div id="chart-area" style="width: 100%; height: calc(100% - 46px); position: relative; overflow: hidden;"></div>
        `;

        const btn = document.getElementById('apply-gauge-btn');
        btn.onmouseover = () => btn.style.background = '#4A5CC4';
        btn.onmouseout = () => btn.style.background = '#5B6FD8';

        document.getElementById('apply-gauge-btn').addEventListener('click', () => {
            const selectedMode = document.querySelector('input[name="gaugeMode"]:checked').value;
            window.currentGaugeMode = selectedMode;
            
            const args = window.gaugeActiveArgs;
            renderViz(args.rawData, args.encodingMap, args.selectedMarksIds, args.styles);
        });
    }

    const chartArea = document.getElementById('chart-area');
    chartArea.innerHTML = '';

    const width = chartArea.offsetWidth || chartArea.clientWidth || 400;
    const height = chartArea.offsetHeight || chartArea.clientHeight || 300;

    const result = await GaugeChart(encodedData, encodingMap, width, height, selectedMarksIds, styles);

    if (result.viz) chartArea.appendChild(result.viz);
    return result;
}

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
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
    unit: "",                               // No unit (KWh fully removed)
    isPercentage: config.isPercentage || false
};

/* ============================================================
   🎚️  GAUGE SIZE TUNING  —  chart छोटा/बड़ा करने के लिए यहीं बदलें
   ------------------------------------------------------------
   sizeBoost : 1.00 = safe fit | 1.18 = बड़ा | 1.30 = और बड़ा
   ============================================================ */
const GAUGE_TUNING = {
    sizeBoost:     1.18,   // ⭐ MAIN KNOB — यही बदलें
    padPct:        0.008,  // किनारों का gap (कम = बड़ा chart)
    topReservePx:  34,     // toggle buttons के लिए ऊपर जगह (0 = और बड़ा)
    labelRadius:   1.15,   // tick labels की दूरी (कम = बड़ा)
    verticalNudge: 0       // -20 = ऊपर खिसकाएँ, +20 = नीचे
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

// ✨ AUTO-FIT: calculates precise font size dynamically to prevent overlap
function fitFontSize(text, maxWidth, baseSize, minSize = 8) {
    const len = String(text).length || 1;
    const approx = maxWidth / (len * 0.56);
    return Math.max(minSize, Math.min(baseSize, approx));
}

// ✨ bbox नापकर container में fit + BOOST + center करता है
function fitVizToContainer(svgNode, width, height, topReserve = 34, pad = 6) {
    if (!svgNode) return;
    const g = svgNode.querySelector('g.gauge-root');
    if (!g) return;

    // pure bbox लेने के लिए पहले transform reset करें
    g.setAttribute('transform', 'translate(0,0) scale(1)');

    let bbox;
    try { bbox = g.getBBox(); } catch (e) { return; }
    if (!bbox || !bbox.width || !bbox.height) return;

    const availW = Math.max(20, width  - pad * 2);
    const availH = Math.max(20, height - pad * 2 - topReserve);

    // ⭐ यहाँ sizeBoost apply होता है
    const k = Math.min(availW / bbox.width, availH / bbox.height) * GAUGE_TUNING.sizeBoost;

    const tx = pad + (availW - bbox.width  * k) / 2 - bbox.x * k;
    const ty = pad + topReserve + (availH - bbox.height * k) / 2 - bbox.y * k
               + GAUGE_TUNING.verticalNudge;

    g.setAttribute('transform', `translate(${tx},${ty}) scale(${k})`);
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

    if (isPctMode) {
        maxScale = 100;
    } else {
        const rawMax = (Math.max(totalValue, totalTarget) * 1.2) || 100;
        const pow10 = Math.pow(10, Math.floor(Math.log10(rawMax)));
        maxScale = Math.ceil(rawMax / (pow10 / 2)) * (pow10 / 2);
    }

    const sampleFormatted = isPctMode ? "100.00%" : formatNumber(maxScale, false, false, true, CONFIG.decimalPlaces);
    const numberLength = sampleFormatted.length;
    const fontScale = Math.min(1.0, Math.max(0.7, 10 / numberLength));

    width = Math.max(width, 100);
    height = Math.max(height, 100);

    const minDim = Math.min(width, height);

    // ---------------- GEOMETRY (tuning-driven) ----------------
    const topReserve = GAUGE_TUNING.topReservePx;
    const pad = Math.max(3, minDim * GAUGE_TUNING.padPct);

    const availW = Math.max(60, width  - pad * 2);
    const availH = Math.max(60, height - pad * 2 - topReserve);

    // approx shape — fitVizToContainer() बाद में exact + boost कर देगा
    const radius = Math.min(availW / 2.65, availH / 2.00);

    const cx = width / 2;
    const cy = pad + topReserve + radius * 1.15;
    // ----------------------------------------------------------

    const svg = d3.create('svg')
        .attr('class', tableau.ClassNameKey.Worksheet)
        .attr('width', '100%')
        .attr('height', '100%')
        .attr('viewBox', `0 0 ${width} ${height}`)
        .attr('preserveAspectRatio', 'xMidYMid meet')
        .style('display', 'block')
        .style('overflow', 'visible')
        .style('background', 'white');

    const chartGroup = svg.append('g')
        .attr('class', 'gauge-root')
        .attr('transform', `translate(${cx}, ${cy})`);

    const startAngle = -Math.PI * 0.75;
    const endAngle = Math.PI * 0.75;
    const totalRange = endAngle - startAngle;

    let valueFraction;
    if (isPctMode) {
        const achievedPct = totalTarget > 0 ? (totalValue / totalTarget) * 100 : 0;
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
        const tickStart = radius * 1.02;
        const tickEnd = radius * 1.09;
        const labelR = radius * GAUGE_TUNING.labelRadius;

        chartGroup.append('line')
            .attr('x1', tickStart * x)
            .attr('y1', tickStart * y)
            .attr('x2', tickEnd * x)
            .attr('y2', tickEnd * y)
            .attr('stroke', '#333')
            .attr('stroke-width', Math.max(1, radius * 0.02));

        let tickStr;
        if (isPctMode) {
            tickStr = formatNumber(f * maxScale, false, true, true, CONFIG.decimalPlaces);
        } else {
            tickStr = formatNumber(f * maxScale, false, false, true, CONFIG.decimalPlaces);
        }

        const isLeft = (i === 0);
        const textAnchor = isLeft ? 'end' : 'start';
        const xOffset = isLeft ? -5 : 5;

        chartGroup.append('text')
            .attr('x', labelR * x + xOffset)
            .attr('y', labelR * y)
            .attr('text-anchor', textAnchor)
            .attr('dominant-baseline', 'middle')
            .style('font-family', 'Arial, sans-serif')
            .style('font-size', fitFontSize(tickStr, radius * 0.85, radius * 0.12) * fontScale + 'px')
            .style('fill', '#333')
            .style('font-weight', 'bold')
            .text(tickStr);
    }

    if (targetKey && totalTarget > 0) {
        const targetPosFraction = isPctMode ? 1.0 : (totalTarget / maxScale);
        const tAngle = startAngle + (Math.min(targetPosFraction, 1) * totalRange);
        chartGroup.append('line')
            .attr('x1', radius * 0.6 * Math.sin(tAngle))
            .attr('y1', radius * 0.6 * -Math.cos(tAngle))
            .attr('x2', radius * 1.12 * Math.sin(tAngle))
            .attr('y2', radius * 1.12 * -Math.cos(tAngle))
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

    // CENTER TEXT VALUE
    const finalCenterValue = isPctMode ? (totalTarget > 0 ? (totalValue / totalTarget) * 100 : 0) : totalValue;
    const centerStr = formatNumber(finalCenterValue, false, isPctMode, false, CONFIG.decimalPlacesForAchievedValue);

    const valueText = chartGroup.append('text')
        .attr('y', radius * 0.20)
        .attr('text-anchor', 'middle')
        .attr('dominant-baseline', 'middle')
        .style('font-family', 'Arial, sans-serif')
        .style('font-size', fitFontSize(centerStr, radius * 1.25, radius * 0.24) * fontScale + 'px')
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
        const currentValue = finalCenterValue * easeProgress;
        valueText.text(formatNumber(currentValue, false, isPctMode, false, CONFIG.decimalPlacesForAchievedValue));
        if (progress < 1) requestAnimationFrame(animateValue);
    }
    requestAnimationFrame(animateValue);

    // % ACHIEVED
    const percentageDisplay = totalTarget > 0 ? (totalValue / totalTarget) * 100 : 0;
    const pctStr = `${percentageDisplay.toFixed(2)}% achieved`;
    const percentageText = chartGroup.append('text')
        .attr('y', radius * 0.46)
        .attr('text-anchor', 'middle')
        .attr('dominant-baseline', 'middle')
        .style('font-family', 'Arial, sans-serif')
        .style('font-size', fitFontSize(pctStr, radius * 1.4, radius * 0.125) * fontScale + 'px')
        .style('font-weight', 'bold')
        .style('fill', '#666')
        .text(pctStr);
    percentageText.raise();

    // TARGET BOTTOM TEXT
    if (targetKey && totalTarget > 0) {
        let targetStr;
        if (isPctMode) {
            targetStr = 'Target: ' + formatNumber(100, false, true, true, CONFIG.decimalPlaces);
        } else {
            targetStr = 'Target: ' + formatNumber(totalTarget, false, false, true, CONFIG.decimalPlaces);
        }

        const targetText = chartGroup.append('text')
            .attr('y', radius * 0.68)
            .attr('text-anchor', 'middle')
            .attr('dominant-baseline', 'middle')
            .style('font-family', 'Arial, sans-serif')
            .style('font-size', fitFontSize(targetStr, radius * 1.5, radius * 0.105) * fontScale + 'px')
            .style('font-weight', '400')
            .style('fill', '#999')
            .text(targetStr);
        targetText.raise();
    }

    const interactionElement = chartGroup.append('circle')
        .attr('r', radius)
        .attr('fill', 'transparent')
        .style('pointer-events', 'none');

    return { viz: svg.node(), interactionElement, allTupleIds, topReserve, pad };
}


// ✨ RENDER VIZ
async function renderViz(rawData, encodingMap, selectedMarksIds, styles) {
    const encodedData = getEncodedData(rawData, encodingMap);
    const content = document.getElementById('content');
    if (!content) {
        console.error('Content div not found!');
        return { viz: null };
    }

    window.gaugeActiveArgs = { rawData, encodingMap, selectedMarksIds, styles };

    content.style.position = 'relative';

    // Global CSS Rule Injection  (id बदला है ताकि पुरानी cached CSS override हो)
    if (!document.getElementById('gauge-global-styles-v2')) {
        const style = document.createElement('style');
        style.id = 'gauge-global-styles-v2';
        style.innerHTML = `
            html, body, #content {
                width: 100% !important;
                height: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
                overflow: hidden !important;
                background-color: white !important;
            }
            #gauge-controls-container {
                position: absolute !important;
                top: 8px !important;
                right: 12px !important;
                z-index: 9999 !important;
                display: flex !important;
                background: #f1f5f9 !important;
                padding: 4px !important;
                border-radius: 30px !important;
                box-shadow: 0 4px 15px rgba(0,0,0,0.08) !important;
                border: 1px solid #e2e8f0 !important;
                font-family: 'Segoe UI', -apple-system, Arial, sans-serif !important;
                user-select: none !important;
                cursor: pointer !important;
            }
            #chart-area {
                width: 100% !important;
                height: 100% !important;
                position: absolute !important;
                top: 0 !important;
                left: 0 !important;
                z-index: 1 !important;
                overflow: visible !important;
            }
            #chart-area > svg {
                display: block !important;
                overflow: visible !important;
            }
        `;
        document.head.appendChild(style);
    }

    let controls = document.getElementById('gauge-controls-container');
    let chartArea = document.getElementById('chart-area');

    if (!controls) {
        content.innerHTML = '';

        controls = document.createElement('div');
        controls.id = 'gauge-controls-container';
        content.appendChild(controls);

        chartArea = document.createElement('div');
        chartArea.id = 'chart-area';
        content.appendChild(chartArea);
    }

    controls.innerHTML = `
        <div id="btn-val-mode" style="padding: 5px 16px; border-radius: 20px; font-size: 12px; font-weight: 700; transition: all 0.25s ease; ${window.currentGaugeMode !== 'percentage' ? 'background: #ffffff; color: #5B6FD8; box-shadow: 0 2px 6px rgba(0,0,0,0.08);' : 'color: #64748b;'}">Raw Number</div>
        <div id="btn-pct-mode" style="padding: 5px 16px; border-radius: 20px; font-size: 12px; font-weight: 700; transition: all 0.25s ease; ${window.currentGaugeMode === 'percentage' ? 'background: #ffffff; color: #5B6FD8; box-shadow: 0 2px 6px rgba(0,0,0,0.08);' : 'color: #64748b;'}">Percentage (%)</div>
    `;

    document.getElementById('btn-val-mode').onclick = () => {
        if (window.currentGaugeMode !== 'value') {
            window.currentGaugeMode = 'value';
            reRenderAll();
        }
    };
    document.getElementById('btn-pct-mode').onclick = () => {
        if (window.currentGaugeMode !== 'percentage') {
            window.currentGaugeMode = 'percentage';
            reRenderAll();
        }
    };

    function reRenderAll() {
        const args = window.gaugeActiveArgs;
        renderViz(args.rawData, args.encodingMap, args.selectedMarksIds, args.styles);
    }

    chartArea.innerHTML = '';

    const rect = chartArea.getBoundingClientRect();
    const width  = Math.round(rect.width)  || chartArea.offsetWidth  || window.innerWidth  || 400;
    const height = Math.round(rect.height) || chartArea.offsetHeight || window.innerHeight || 300;

    const result = await GaugeChart(encodedData, encodingMap, width, height, selectedMarksIds, styles);

    if (result.viz) {
        chartArea.appendChild(result.viz);

        const tR = (result.topReserve != null) ? result.topReserve : GAUGE_TUNING.topReservePx;
        const pd = (result.pad != null) ? result.pad : 6;

        fitVizToContainer(result.viz, width, height, tR, pd);
        requestAnimationFrame(() => fitVizToContainer(result.viz, width, height, tR, pd));
        setTimeout(() => fitVizToContainer(result.viz, width, height, tR, pd), 60);
    }
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
        const doResize = () => {
            clearTimeout(resizeTimeout);
            resizeTimeout = setTimeout(() => {
                const args = window.gaugeActiveArgs;
                if (args) {
                    renderViz(args.rawData, args.encodingMap, args.selectedMarksIds, args.styles);
                }
            }, 120);
        };

        window.addEventListener('resize', doResize);

        const contentEl = document.getElementById('content');
        if (contentEl && window.ResizeObserver) {
            const ro = new ResizeObserver(doResize);
            ro.observe(contentEl);
        }

        document.body.addEventListener('click', async (e) => {
            if (e.target.closest('#gauge-controls-container')) return;

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
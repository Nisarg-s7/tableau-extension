/* global d3 */
/* global tinycolor */
/* global tableau */

const backgroundColor = tinycolor('white');
const palette = ['#5B6FD8', '#D3D3D3', '#4e79a7', '#f28e2c'];

// 👇 YAHAN SE APNI SETTINGS BADALEN 👇
// Tableau se aane wale config ko parse karke default values set karein
let config = {};
try {
    if (styles && styles.configJson) {
        config = JSON.parse(styles.configJson);
    }
} catch (e) {
    console.error("Error parsing config:", e);
}

// Agar Tableau se config nahi milta, toh ye default values istemal hongi
const CONFIG = {
    measure: config.measure || "Sales",
    format: config.format || "#,##0.00",
    useSuffix: config.useSuffix !== false,
    decimalPlaces: config.decimalPlaces || 1,
    prefix: config.prefix || "",
    unit: config.unit || " KWh",
    
    // ✨ NAYA OPTION YAHAN ADD KIYA HAI ✨
    // Agar aapka data Percentage hai (0 se 100 ke beech), toh ise TRUE karein.
    // Agar aapka data Rupees/Dollar/KWh jaisa hai, toh ise FALSE rakhein.
    isPercentage: config.isPercentage || false 
};
// 👆 YAHAN SE SETTINGS KHATAM 👆

// ✨ UPDATED FUNCTION ✨
// Yeh function ab 'isPercentage' ke hisaab se format decide karega
function formatNumber(value, isPercentage) {
    // Agar isPercentage TRUE hai, toh sirf % lagao
    if (isPercentage) {
        return value.toFixed(CONFIG.decimalPlaces) + "%";
    }

    // Agar isPercentage FALSE hai, toh K/M suffix lagao (Purana logic)
    if (CONFIG.useSuffix) {
        if (value >= 1000000) {
            return (value / 1000000).toFixed(CONFIG.decimalPlaces) + "M";
        } else if (value >= 1000) {
            return (value / 1000).toFixed(CONFIG.decimalPlaces) + "K";
        }
    }
    
    // Agar koi suffix nahi, toh sirf unit (KWh, $, etc.) laga do
    return Math.round(value) + CONFIG.unit;
}

// MAIN GAUGE CHART FUNCTION
async function GaugeChart(encodedData, encodingMap, width, height, selectedMarksIds, styles) {
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
                totalTarget += parseFloat(row[targetKey]?.[0]?.value || 0);
            }
            if (row.tupleId) allTupleIds.push(row.tupleId);
        });

        // ✨ SMART LOGIC YAHAN HAI ✨
        // Yeh automatically detect karega ki scale kya honi chahiye
        if (CONFIG.isPercentage) {
            // Agar Percentage hai, to scale hamesha 100 hogi
            maxScale = 100;
        } else {
            // Agar Number hai (KWh/Rupees), to value ke aadhar par scale banegi
            if (totalTarget > 0) {
                const rawMax = totalTarget * 1.25;
                const pow10 = Math.pow(10, Math.floor(Math.log10(rawMax)));
                maxScale = Math.ceil(rawMax / (pow10 / 2)) * (pow10 / 2);
            } else {
                const rawMax = totalValue * 1.4;
                const pow10 = Math.pow(10, Math.floor(Math.log10(rawMax)));
                maxScale = Math.ceil(rawMax / (pow10 / 2)) * (pow10 / 2);
            }
        }

        window._lockedFinalValues = {
            totalValue,
            totalTarget,
            maxScale,
            allTupleIds
        };
    }

    width = Math.max(width, 100);
    height = Math.max(height, 100);

    const minDim = Math.min(width, height);
    const margin = { 
        top: minDim * 0.12, 
        right: minDim * 0.15, 
        bottom: minDim * 0.12, 
        left: minDim * 0.15 
    };
    
    const cx = width / 2;
    const cy = height / 2;
    const radius = Math.min(
        (width - margin.left - margin.right) / 2.6,
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
    const valueFraction = Math.min(Math.max(totalValue / maxScale, 0), 1);
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
            // ✨ YAHAN BHI UPDATE ✨
            .text(formatNumber(f * maxScale, CONFIG.isPercentage));
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

    // Animated Value
    const valueText = chartGroup.append('text')
        .attr('y', radius * 0.30)
        .attr('text-anchor', 'middle')
        .attr('dominant-baseline', 'middle')
        .style('font-family', 'Arial, sans-serif')
        .style('font-size', Math.max(16, radius * 0.25) + 'px')
        .style('font-weight', 'bold')
        .style('fill', palette[0])
        // ✨ YAHAN BHI UPDATE ✨
        .text(CONFIG.prefix + formatNumber(totalValue, CONFIG.isPercentage) + CONFIG.unit);
    valueText.raise();

    const animationDuration = 1000;
    const startTime = Date.now();
    
    function animateValue() {
        const elapsed = Date.now() - startTime;
        const progress = Math.min(elapsed / animationDuration, 1);
        const easeProgress = 1 - Math.pow(1 - progress, 3);
        const currentValue = totalValue * easeProgress;
        
        valueText.text(CONFIG.prefix + formatNumber(totalValue, CONFIG.isPercentage) + CONFIG.unit);

        if (progress < 1) requestAnimationFrame(animateValue);
    }
    requestAnimationFrame(animateValue);

    // Percentage Text (Yeh hamesha percentage dikhayega achievement ke liye)
    let percentageDisplay = 0;
    if (totalTarget > 0) {
        const actualPercentage = (totalValue / totalTarget) * 100;
        percentageDisplay = Math.min(actualPercentage, 100);
    } else {
        percentageDisplay = totalValue > 0 ? 100 : 0;
    }
    const percentageText = chartGroup.append('text')
        .attr('y', radius * 0.70)
        .attr('text-anchor', 'middle')
        .attr('dominant-baseline', 'middle')
        .style('font-family', 'Arial, sans-serif')
        .style('font-size', Math.max(10, radius * 0.15) + 'px')
        .style('font-weight', 'bold')
        .style('fill', '#666')
        .text(`${percentageDisplay.toFixed(1)}% achieved`);
    percentageText.raise();

    // Target Text
    if (targetKey && totalTarget > 0) {
        const targetText = chartGroup.append('text')
            .attr('y', radius * 0.85)
            .attr('text-anchor', 'middle')
            .attr('dominant-baseline', 'middle')
            .style('font-family', 'Arial, sans-serif')
            .style('font-size', Math.max(8, radius * 0.12) + 'px')
            .style('font-weight', '400')
            .style('fill', '#999')
            // ✨ YAHAN BHI UPDATE ✨
            .text('Target: ' + formatNumber(totalTarget, CONFIG.isPercentage));

        targetText.raise();
    }

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

// HELPERS (Inmein koi change nahi hai)
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
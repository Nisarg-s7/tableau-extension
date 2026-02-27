// =============================================
// BAR CHART WITH SIZE MARK CARD - FIXED VERSION
// =============================================

console.log('📦 Script loading...');

// =============================================
// CONFIGURATION
// =============================================
const CONFIG = {
    barRatio: 0.75,      // Default bar width (75%)
    chartScale: 1.0,     // Default scale (100%)
    sizeImpact: 0.3,     // How much size affects width
    colors: ['#4e79a7', '#f28e2c', '#e15759', '#76b7b2', '#59a14f', '#edc949']
};

// Global variables
let chartData = [];
let worksheet = null;
let isTableauReady = false;
let livePreview = true; // Enable live preview when sliding

// =============================================
// INITIALIZATION
// =============================================
document.addEventListener('DOMContentLoaded', function() {
    console.log('✅ DOM ready');
    
    // Setup buttons
    document.getElementById('config-btn').addEventListener('click', openSizeConfig);
    document.getElementById('refresh-btn').addEventListener('click', refreshChart);
    
    // Create Size Mark Card
    createSizeMarkCard();
    
    // Initialize Tableau or use test data
    initTableau();
});

// =============================================
// SIZE MARK CARD UI - FIXED
// =============================================
function createSizeMarkCard() {
    console.log('🎨 Creating Size Mark Card...');
    
    // Remove if exists
    const existing = document.getElementById('size-mark-card');
    if (existing) existing.remove();
    
    // Create card HTML
    const cardHTML = `
        <div id="size-mark-card">
            <div class="mark-header">
                <span style="font-weight: bold;">📏 Size Control</span>
                <div class="header-controls">
                    <label class="live-toggle">
                        <input type="checkbox" id="live-preview-toggle" checked>
                        <span>Live</span>
                    </label>
                    <span class="status-badge" id="size-status">Ready</span>
                </div>
            </div>
            
            <div class="mark-body">
                <!-- Chart Scale Control -->
                <div class="control-section highlight">
                    <div class="control-label">
                        <span>🔍 Chart Scale</span>
                        <span class="control-value" id="scale-display">100%</span>
                    </div>
                    <input type="range" class="size-slider" id="scale-slider" 
                           min="50" max="200" value="100" step="5">
                    <div class="control-labels">
                        <span>50%</span>
                        <span>100%</span>
                        <span>200%</span>
                    </div>
                    <div class="quick-buttons" id="scale-quick-buttons">
                        <button class="quick-btn" data-value="50">50%</button>
                        <button class="quick-btn" data-value="75">75%</button>
                        <button class="quick-btn active" data-value="100">100%</button>
                        <button class="quick-btn" data-value="150">150%</button>
                        <button class="quick-btn" data-value="200">200%</button>
                    </div>
                </div>
                
                <!-- Chart Thickness Control -->
                <div class="control-section special">
                    <div class="control-label">
                        <span>📊 Bar Thickness</span>
                        <span class="control-value green" id="width-display">75%</span>
                    </div>
                    <input type="range" class="size-slider green" id="width-slider" 
                           min="10" max="100" value="75" step="1">
                    <div class="control-labels">
                        <span>10%</span>
                        <span>55%</span>
                        <span>100%</span>
                    </div>
                    <div class="quick-buttons" id="width-quick-buttons">
                        <button class="quick-btn" data-value="25">Thin</button>
                        <button class="quick-btn" data-value="50">Medium</button>
                        <button class="quick-btn active" data-value="75">Normal</button>
                        <button class="quick-btn" data-value="90">Thick</button>
                    </div>
                    <p class="hint">💡 Controls thickness of ALL bars</p>
                </div>
                
                <!-- Size Impact Control -->
                <div class="control-section">
                    <div class="control-label">
                        <span>🎚️ Size Variation</span>
                        <span class="control-value purple" id="impact-display">30%</span>
                    </div>
                    <input type="range" class="size-slider purple" id="impact-slider" 
                           min="0" max="45" value="30" step="5">
                    <div class="control-labels">
                        <span>0%</span>
                        <span>Uniform</span>
                        <span>45%</span>
                    </div>
                    <p class="hint">💡 Higher = more width variation based on size data</p>
                </div>
            </div>
            
            <div class="mark-footer">
                <button class="mark-btn reset" id="reset-size-btn">🔄 Reset</button>
                <button class="mark-btn apply" id="apply-size-btn">✓ Apply</button>
            </div>
        </div>
    `;
    
    // Create styles
    const styleHTML = `
        <style id="size-mark-card-style">
            #size-mark-card {
                position: fixed;
                top: 10px;
                right: 10px;
                width: 300px;
                background: white;
                border: 2px solid #4e79a7;
                border-radius: 8px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                font-family: Arial, sans-serif;
                font-size: 12px;
                z-index: 9999;
                animation: slideInRight 0.3s;
                user-select: none;
            }
            
            @keyframes slideInRight {
                from { transform: translateX(100%); opacity: 0; }
                to { transform: translateX(0); opacity: 1; }
            }
            
            .mark-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 10px 12px;
                background: linear-gradient(135deg, #4e79a7, #76b7b2);
                color: white;
                border-radius: 6px 6px 0 0;
                font-size: 13px;
            }
            
            .header-controls {
                display: flex;
                align-items: center;
                gap: 8px;
            }
            
            .live-toggle {
                display: flex;
                align-items: center;
                gap: 4px;
                font-size: 10px;
                cursor: pointer;
            }
            
            .live-toggle input {
                width: 14px;
                height: 14px;
                cursor: pointer;
            }
            
            .status-badge {
                background: rgba(255,255,255,0.3);
                padding: 2px 8px;
                border-radius: 10px;
                font-size: 10px;
                font-weight: bold;
                transition: all 0.3s;
            }
            
            .status-badge.success {
                background: #4caf50;
            }
            
            .mark-body {
                padding: 12px;
                max-height: calc(100vh - 200px);
                overflow-y: auto;
            }
            
            .control-section {
                margin-bottom: 15px;
                padding: 12px;
                background: #f8f9fa;
                border-radius: 6px;
                border: 1px solid #e0e0e0;
                transition: all 0.2s;
            }
            
            .control-section:hover {
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            }
            
            .control-section.highlight {
                background: linear-gradient(135deg, #e3f2fd, #bbdefb);
                border: 2px solid #2196f3;
            }
            
            .control-section.special {
                background: linear-gradient(135deg, #e8f5e9, #c8e6c9);
                border: 2px solid #4caf50;
            }
            
            .control-label {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 8px;
                font-weight: 600;
            }
            
            .control-value {
                font-size: 22px;
                font-weight: 900;
                color: #2196f3;
                min-width: 60px;
                text-align: right;
                transition: transform 0.1s;
            }
            
            .control-value.pulse {
                transform: scale(1.1);
            }
            
            .control-value.green { color: #4caf50; }
            .control-value.purple { color: #9c27b0; }
            
            .size-slider {
                width: 100%;
                height: 8px;
                -webkit-appearance: none;
                appearance: none;
                background: linear-gradient(to right, #ddd 0%, #2196f3 50%, #ddd 100%);
                background-size: 100% 100%;
                border-radius: 4px;
                outline: none;
                margin: 10px 0;
                cursor: pointer;
            }
            
            .size-slider::-webkit-slider-thumb {
                -webkit-appearance: none;
                appearance: none;
                width: 20px;
                height: 20px;
                background: #2196f3;
                border-radius: 50%;
                cursor: grab;
                border: 3px solid white;
                box-shadow: 0 2px 6px rgba(0,0,0,0.3);
                transition: transform 0.1s, box-shadow 0.1s;
            }
            
            .size-slider::-webkit-slider-thumb:hover {
                transform: scale(1.15);
                box-shadow: 0 3px 10px rgba(0,0,0,0.4);
            }
            
            .size-slider::-webkit-slider-thumb:active {
                cursor: grabbing;
                transform: scale(1.2);
            }
            
            .size-slider::-moz-range-thumb {
                width: 20px;
                height: 20px;
                background: #2196f3;
                border-radius: 50%;
                cursor: grab;
                border: 3px solid white;
                box-shadow: 0 2px 6px rgba(0,0,0,0.3);
            }
            
            .size-slider.green { background: linear-gradient(to right, #ddd 0%, #4caf50 50%, #ddd 100%); }
            .size-slider.green::-webkit-slider-thumb { background: #4caf50; }
            .size-slider.green::-moz-range-thumb { background: #4caf50; }
            
            .size-slider.purple { background: linear-gradient(to right, #ddd 0%, #9c27b0 50%, #ddd 100%); }
            .size-slider.purple::-webkit-slider-thumb { background: #9c27b0; }
            .size-slider.purple::-moz-range-thumb { background: #9c27b0; }
            
            .control-labels {
                display: flex;
                justify-content: space-between;
                font-size: 10px;
                color: #888;
                margin-top: 2px;
            }
            
            .quick-buttons {
                display: flex;
                gap: 4px;
                margin-top: 10px;
                flex-wrap: wrap;
            }
            
            .quick-btn {
                flex: 1;
                min-width: 45px;
                padding: 6px 4px;
                border: 2px solid #2196f3;
                background: white;
                color: #2196f3;
                border-radius: 4px;
                cursor: pointer;
                font-weight: bold;
                font-size: 10px;
                transition: all 0.2s;
            }
            
            .quick-btn:hover {
                background: #e3f2fd;
                transform: translateY(-2px);
            }
            
            .quick-btn.active {
                background: #2196f3;
                color: white;
            }
            
            #width-quick-buttons .quick-btn {
                border-color: #4caf50;
                color: #4caf50;
            }
            
            #width-quick-buttons .quick-btn:hover {
                background: #e8f5e9;
            }
            
            #width-quick-buttons .quick-btn.active {
                background: #4caf50;
                color: white;
            }
            
            .mark-footer {
                display: flex;
                gap: 8px;
                padding: 10px 12px;
                background: #f5f5f5;
                border-radius: 0 0 6px 6px;
                border-top: 1px solid #e0e0e0;
            }
            
            .mark-btn {
                flex: 1;
                padding: 10px;
                border: none;
                border-radius: 5px;
                cursor: pointer;
                font-weight: bold;
                font-size: 13px;
                transition: all 0.2s;
            }
            
            .mark-btn.apply {
                background: linear-gradient(135deg, #4caf50, #66bb6a);
                color: white;
                box-shadow: 0 2px 6px rgba(76,175,80,0.3);
            }
            
            .mark-btn.apply:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(76,175,80,0.4);
            }
            
            .mark-btn.reset {
                background: white;
                color: #e65100;
                border: 2px solid #ff9800;
            }
            
            .mark-btn.reset:hover {
                background: #fff3e0;
                transform: translateY(-2px);
            }
            
            .hint {
                margin-top: 8px;
                padding: 6px 8px;
                background: rgba(255,255,255,0.8);
                border-radius: 4px;
                font-size: 10px;
                color: #555;
                font-style: italic;
                line-height: 1.4;
                border-left: 3px solid #ff9800;
            }
        </style>
    `;
    
    // Add to page
    document.head.insertAdjacentHTML('beforeend', styleHTML);
    document.body.insertAdjacentHTML('beforeend', cardHTML);
    
    // Setup event listeners (FIXED)
    setupSizeCardListeners();
    
    console.log('✅ Size Mark Card created');
}

// =============================================
// FIXED: SETUP SIZE CARD LISTENERS
// =============================================
function setupSizeCardListeners() {
    // Get elements
    const scaleSlider = document.getElementById('scale-slider');
    const widthSlider = document.getElementById('width-slider');
    const impactSlider = document.getElementById('impact-slider');
    const liveToggle = document.getElementById('live-preview-toggle');
    const resetBtn = document.getElementById('reset-size-btn');
    const applyBtn = document.getElementById('apply-size-btn');
    
    // Scale slider events
    scaleSlider.addEventListener('input', function() {
        const value = parseInt(this.value);
        updateSliderDisplay('scale', value);
        updateScaleQuickButtons(value);
        if (livePreview) applyLivePreview();
    });
    
    scaleSlider.addEventListener('change', function() {
        if (!livePreview) applyLivePreview();
    });
    
    // Width slider events
    widthSlider.addEventListener('input', function() {
        const value = parseInt(this.value);
        updateSliderDisplay('width', value);
        updateWidthQuickButtons(value);
        if (livePreview) applyLivePreview();
    });
    
    widthSlider.addEventListener('change', function() {
        if (!livePreview) applyLivePreview();
    });
    
    // Impact slider events
    impactSlider.addEventListener('input', function() {
        const value = parseInt(this.value);
        updateSliderDisplay('impact', value);
        if (livePreview) applyLivePreview();
    });
    
    impactSlider.addEventListener('change', function() {
        if (!livePreview) applyLivePreview();
    });
    
    // Live preview toggle
    liveToggle.addEventListener('change', function() {
        livePreview = this.checked;
        console.log('Live preview:', livePreview);
    });
    
    // Quick buttons for scale
    document.querySelectorAll('#scale-quick-buttons .quick-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const value = parseInt(this.dataset.value);
            scaleSlider.value = value;
            updateSliderDisplay('scale', value);
            updateScaleQuickButtons(value);
            applyLivePreview();
        });
    });
    
    // Quick buttons for width
    document.querySelectorAll('#width-quick-buttons .quick-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const value = parseInt(this.dataset.value);
            widthSlider.value = value;
            updateSliderDisplay('width', value);
            updateWidthQuickButtons(value);
            applyLivePreview();
        });
    });
    
    // Reset button
    resetBtn.addEventListener('click', resetSizeCard);
    
    // Apply button
    applyBtn.addEventListener('click', applySizeCard);
    
    // Initialize values from CONFIG
    syncSlidersFromConfig();
}

// =============================================
// HELPER FUNCTIONS FOR SLIDERS
// =============================================
function updateSliderDisplay(type, value) {
    const display = document.getElementById(`${type}-display`);
    if (display) {
        display.textContent = value + '%';
        // Pulse animation
        display.classList.add('pulse');
        setTimeout(() => display.classList.remove('pulse'), 100);
    }
}

function updateScaleQuickButtons(value) {
    document.querySelectorAll('#scale-quick-buttons .quick-btn').forEach(btn => {
        const btnValue = parseInt(btn.dataset.value);
        btn.classList.toggle('active', btnValue === value);
    });
}

function updateWidthQuickButtons(value) {
    document.querySelectorAll('#width-quick-buttons .quick-btn').forEach(btn => {
        const btnValue = parseInt(btn.dataset.value);
        // Approximate match for width buttons
        const isActive = Math.abs(btnValue - value) <= 5;
        btn.classList.toggle('active', isActive);
    });
}

function syncSlidersFromConfig() {
    const scaleSlider = document.getElementById('scale-slider');
    const widthSlider = document.getElementById('width-slider');
    const impactSlider = document.getElementById('impact-slider');
    
    if (scaleSlider) {
        const scaleValue = Math.round(CONFIG.chartScale * 100);
        scaleSlider.value = scaleValue;
        updateSliderDisplay('scale', scaleValue);
        updateScaleQuickButtons(scaleValue);
    }
    
    if (widthSlider) {
        const widthValue = Math.round(CONFIG.barRatio * 100);
        widthSlider.value = widthValue;
        updateSliderDisplay('width', widthValue);
        updateWidthQuickButtons(widthValue);
    }
    
    if (impactSlider) {
        const impactValue = Math.round(CONFIG.sizeImpact * 100);
        impactSlider.value = impactValue;
        updateSliderDisplay('impact', impactValue);
    }
}

function applyLivePreview() {
    // Get current slider values
    const scaleValue = parseInt(document.getElementById('scale-slider').value);
    const widthValue = parseInt(document.getElementById('width-slider').value);
    const impactValue = parseInt(document.getElementById('impact-slider').value);
    
    // Update CONFIG
    CONFIG.chartScale = scaleValue / 100;
    CONFIG.barRatio = widthValue / 100;
    CONFIG.sizeImpact = impactValue / 100;
    
    // Re-render chart
    renderChart();
    
    // Update status briefly
    const status = document.getElementById('size-status');
    status.textContent = 'Updated';
    status.classList.add('success');
    
    setTimeout(() => {
        status.textContent = 'Ready';
        status.classList.remove('success');
    }, 500);
}

function resetSizeCard() {
    console.log('🔄 Resetting Size Card...');
    
    // Reset to defaults
    CONFIG.chartScale = 1.0;
    CONFIG.barRatio = 0.75;
    CONFIG.sizeImpact = 0.3;
    
    // Sync sliders
    syncSlidersFromConfig();
    
    // Re-render
    renderChart();
    
    // Status update
    const status = document.getElementById('size-status');
    status.textContent = 'Reset!';
    status.style.background = '#ff9800';
    
    setTimeout(() => {
        status.textContent = 'Ready';
        status.style.background = 'rgba(255,255,255,0.3)';
    }, 1500);
}

function applySizeCard() {
    console.log('✅ Applying Size Card settings...');
    
    // Get values from sliders
    CONFIG.chartScale = parseInt(document.getElementById('scale-slider').value) / 100;
    CONFIG.barRatio = parseInt(document.getElementById('width-slider').value) / 100;
    CONFIG.sizeImpact = parseInt(document.getElementById('impact-slider').value) / 100;
    
    console.log('Applied config:', CONFIG);
    
    // Update status
    const status = document.getElementById('size-status');
    status.textContent = 'Applied!';
    status.classList.add('success');
    
    setTimeout(() => {
        status.textContent = 'Ready';
        status.classList.remove('success');
    }, 2000);
    
    // Save settings
    saveSettings();
    
    // Re-render chart
    renderChart();
    
    // Update main status
    updateStatus(`Scale: ${Math.round(CONFIG.chartScale * 100)}%, Width: ${Math.round(CONFIG.barRatio * 100)}%`);
}

// =============================================
// TABLEAU INITIALIZATION
// =============================================
async function initTableau() {
    updateStatus('Connecting to Tableau...');
    
    try {
        if (typeof tableau === 'undefined' || !tableau.extensions) {
            console.log('⚠️ Tableau API not found, using test data');
            useTestData();
            return;
        }
        
        await tableau.extensions.initializeAsync();
        console.log('✅ Tableau connected');
        isTableauReady = true;
        
        loadSettings();
        syncSlidersFromConfig();
        
        const dashboard = tableau.extensions.dashboardContent.dashboard;
        console.log('📊 Dashboard:', dashboard.name);
        
        if (dashboard.worksheets.length === 0) {
            showError('No worksheets in dashboard!');
            return;
        }
        
        worksheet = dashboard.worksheets[0];
        console.log('📈 Worksheet:', worksheet.name);
        
        await fetchTableauData();
        
        worksheet.addEventListener(tableau.TableauEventType.FilterChanged, fetchTableauData);
        worksheet.addEventListener(tableau.TableauEventType.MarkSelectionChanged, fetchTableauData);
        
    } catch (err) {
        console.error('❌ Tableau error:', err);
        useTestData();
    }
}

// =============================================
// DATA HANDLING
// =============================================
async function fetchTableauData() {
    updateStatus('Fetching data...');
    
    try {
        const dataTable = await worksheet.getSummaryDataAsync();
        console.log('📊 Got', dataTable.totalRowCount, 'rows');
        
        chartData = processTableauData(dataTable);
        console.log('✅ Processed', chartData.length, 'items');
        
        renderChart();
        updateStatus('Ready');
        
    } catch (err) {
        console.error('❌ Data error:', err);
        showError('Failed to get data: ' + err.message);
    }
}

function processTableauData(dataTable) {
    const cols = dataTable.columns;
    const rows = dataTable.data;
    
    let catIdx = 0;
    let valIdx = 1;
    let sizeIdx = -1;
    
    cols.forEach((col, i) => {
        if (col.dataType === 'string' && catIdx === 0) catIdx = i;
        if ((col.dataType === 'float' || col.dataType === 'int') && valIdx === 1) valIdx = i;
        if ((col.dataType === 'float' || col.dataType === 'int') && i !== valIdx) sizeIdx = i;
    });
    
    const map = new Map();
    
    rows.forEach(row => {
        const cat = row[catIdx].formattedValue || 'Unknown';
        const val = parseFloat(row[valIdx].value) || 0;
        const size = sizeIdx >= 0 ? parseFloat(row[sizeIdx].value) : null;
        
        if (map.has(cat)) {
            map.get(cat).value += val;
            if (size !== null) map.get(cat).size += size;
        } else {
            map.set(cat, { category: cat, value: val, size: size });
        }
    });
    
    const result = Array.from(map.values());
    result.sort((a, b) => b.value - a.value);
    
    return result;
}

function useTestData() {
    console.log('📋 Using test data');
    updateStatus('Test mode');
    
    chartData = [
        { category: 'Technology', value: 836154, size: 120 },
        { category: 'Furniture', value: 741999, size: 85 },
        { category: 'Office Supplies', value: 719047, size: 150 },
        { category: 'Phones', value: 330007, size: 60 },
        { category: 'Chairs', value: 328449, size: 90 },
        { category: 'Storage', value: 223843, size: 45 },
        { category: 'Tables', value: 206965, size: 70 },
        { category: 'Copiers', value: 149528, size: 30 }
    ];
    
    hideLoading();
    renderChart();
}

// =============================================
// CHART RENDERING
// =============================================
function renderChart() {
    console.log('🎨 Rendering chart - Scale:', CONFIG.chartScale, 'Width:', CONFIG.barRatio);
    
    hideLoading();
    
    if (chartData.length === 0) {
        document.getElementById('chart').innerHTML = '';
        showError('No data to display');
        return;
    }
    
    const container = document.getElementById('chart-area');
    const width = container.clientWidth;
    const height = container.clientHeight;
    
    const margin = { top: 20, right: 30, bottom: 50, left: 80 };
    const innerW = width - margin.left - margin.right;
    const innerH = height - margin.top - margin.bottom;
    
    d3.select('#chart').selectAll('*').remove();
    
    const svg = d3.select('#chart')
        .attr('width', width)
        .attr('height', height)
        .style('transform', `scale(${CONFIG.chartScale})`)
        .style('transform-origin', 'top left');
    
    const g = svg.append('g')
        .attr('transform', `translate(${margin.left},${margin.top})`);
    
    // X Scale
    const x = d3.scaleBand()
        .domain(chartData.map(d => d.category))
        .range([0, innerW])
        .padding(0.1);
    
    // Y Scale
    const y = d3.scaleLinear()
        .domain([0, d3.max(chartData, d => d.value) * 1.1])
        .nice()
        .range([innerH, 0]);
    
    // Color scale
    const color = d3.scaleOrdinal()
        .domain(chartData.map(d => d.category))
        .range(CONFIG.colors);
    
    // Size scale
    let sizeScale = null;
    const hasSizeData = chartData.some(d => d.size !== null);
    
    if (hasSizeData && CONFIG.sizeImpact > 0) {
        const sizes = chartData.filter(d => d.size !== null).map(d => d.size);
        const minRatio = Math.max(0.15, CONFIG.barRatio - CONFIG.sizeImpact);
        const maxRatio = Math.min(0.98, CONFIG.barRatio + CONFIG.sizeImpact);
        
        sizeScale = d3.scaleLinear()
            .domain([d3.min(sizes), d3.max(sizes)])
            .range([minRatio, maxRatio]);
    }
    
    // Bar width function
    const getBarWidth = (d) => {
        let ratio = CONFIG.barRatio;
        if (sizeScale && d.size !== null) {
            ratio = sizeScale(d.size);
        }
        return x.bandwidth() * ratio;
    };
    
    // Grid lines
    g.append('g')
        .attr('class', 'grid')
        .call(d3.axisLeft(y).tickSize(-innerW).tickFormat(''))
        .selectAll('line')
        .style('stroke', '#e0e0e0')
        .style('stroke-dasharray', '3,3');
    
    g.selectAll('.grid .domain').remove();
    
    // Bars with animation
    g.selectAll('.bar')
        .data(chartData)
        .enter()
        .append('rect')
        .attr('class', 'bar')
        .attr('x', d => x(d.category) + (x.bandwidth() - getBarWidth(d)) / 2)
        .attr('y', innerH)
        .attr('width', d => getBarWidth(d))
        .attr('height', 0)
        .attr('fill', d => color(d.category))
        .attr('rx', 4)
        .style('cursor', 'pointer')
        .transition()
        .duration(300)
        .attr('y', d => y(d.value))
        .attr('height', d => innerH - y(d.value));
    
    // Mouse events
    g.selectAll('.bar')
        .on('mouseenter', function(event, d) {
            d3.select(this)
                .transition()
                .duration(100)
                .attr('opacity', 0.7)
                .attr('stroke', '#333')
                .attr('stroke-width', 2);
            showTooltip(event, d);
        })
        .on('mouseleave', function() {
            d3.select(this)
                .transition()
                .duration(100)
                .attr('opacity', 1)
                .attr('stroke', 'none');
            hideTooltip();
        });
    
    // Value labels
    g.selectAll('.label')
        .data(chartData)
        .enter()
        .append('text')
        .attr('class', 'label')
        .attr('x', d => x(d.category) + x.bandwidth() / 2)
        .attr('y', d => y(d.value) - 8)
        .attr('text-anchor', 'middle')
        .style('font-size', '11px')
        .style('font-weight', 'bold')
        .style('fill', '#333')
        .style('opacity', 0)
        .text(d => formatNum(d.value))
        .transition()
        .delay(300)
        .duration(200)
        .style('opacity', 1);
    
    // X Axis
    const xAxis = g.append('g')
        .attr('transform', `translate(0,${innerH})`)
        .call(d3.axisBottom(x).tickSize(0));
    
    if (chartData.length > 5 || x.bandwidth() < 80) {
        xAxis.selectAll('text')
            .attr('transform', 'rotate(-35)')
            .style('text-anchor', 'end')
            .attr('dx', '-5px')
            .attr('dy', '5px');
    }
    
    xAxis.selectAll('text').style('font-size', '11px');
    xAxis.select('.domain').style('stroke', '#ccc');
    
    // Y Axis
    const yAxis = g.append('g')
        .call(d3.axisLeft(y).ticks(5).tickFormat(formatNum));
    
    yAxis.selectAll('text').style('font-size', '11px');
    yAxis.select('.domain').style('stroke', '#ccc');
    
    console.log('✅ Chart rendered');
}

function formatNum(n) {
    if (n >= 1000000) return (n / 1000000).toFixed(1) + 'M';
    if (n >= 1000) return (n / 1000).toFixed(1) + 'K';
    return Math.round(n).toString();
}

// =============================================
// TOOLTIP
// =============================================
function showTooltip(event, d) {
    let tip = document.getElementById('tooltip');
    if (!tip) {
        tip = document.createElement('div');
        tip.id = 'tooltip';
        tip.style.cssText = `
            position: fixed;
            background: rgba(0,0,0,0.9);
            color: white;
            padding: 12px 16px;
            border-radius: 8px;
            font-size: 13px;
            pointer-events: none;
            z-index: 99999;
            box-shadow: 0 4px 15px rgba(0,0,0,0.3);
            max-width: 200px;
        `;
        document.body.appendChild(tip);
    }
    
    tip.innerHTML = `
        <strong style="font-size: 14px;">${d.category}</strong><br>
        <span style="color: #90caf9;">Value:</span> ${formatNum(d.value)}
        ${d.size !== null ? `<br><span style="color: #a5d6a7;">Size:</span> ${formatNum(d.size)}` : ''}
    `;
    tip.style.display = 'block';
    tip.style.left = (event.pageX + 15) + 'px';
    tip.style.top = (event.pageY - 10) + 'px';
}

function hideTooltip() {
    const tip = document.getElementById('tooltip');
    if (tip) tip.style.display = 'none';
}

// =============================================
// CONFIG MODAL (Alternative to Size Card)
// =============================================
function openSizeConfig() {
    console.log('📏 Opening size config modal...');
    closeModal();
    
    const modalHTML = `
        <div id="modal-overlay">
            <div id="modal-box">
                <div id="modal-header">
                    <span>📏 Size Configuration</span>
                    <span id="modal-close">×</span>
                </div>
                
                <div id="modal-body">
                    <div class="config-group highlight">
                        <label>🔍 Chart Scale (Zoom)</label>
                        <div class="big-value" id="modal-scale-val">${Math.round(CONFIG.chartScale * 100)}%</div>
                        <input type="range" id="modal-scale-range" min="50" max="200" 
                               value="${Math.round(CONFIG.chartScale * 100)}">
                        <div class="range-labels">
                            <span>50%</span><span>100%</span><span>200%</span>
                        </div>
                    </div>
                    
                    <div class="config-group">
                        <label>📊 Bar Width</label>
                        <div class="big-value green" id="modal-width-val">${Math.round(CONFIG.barRatio * 100)}%</div>
                        <input type="range" id="modal-width-range" min="20" max="98" 
                               value="${Math.round(CONFIG.barRatio * 100)}">
                        <div class="range-labels">
                            <span>Thin</span><span>Thick</span>
                        </div>
                    </div>
                    
                    <div class="config-group">
                        <label>🎚️ Size Field Impact</label>
                        <div class="big-value purple" id="modal-impact-val">${Math.round(CONFIG.sizeImpact * 100)}%</div>
                        <input type="range" id="modal-impact-range" min="0" max="45" 
                               value="${Math.round(CONFIG.sizeImpact * 100)}">
                        <div class="range-labels">
                            <span>Uniform</span><span>Varied</span>
                        </div>
                    </div>
                </div>
                
                <div id="modal-footer">
                    <button class="btn-reset" id="modal-reset-btn">🔄 Reset</button>
                    <button class="btn-apply" id="modal-apply-btn">✓ Apply</button>
                    <button class="btn-cancel" id="modal-cancel-btn">Cancel</button>
                </div>
            </div>
        </div>
    `;
    
    const styleHTML = `
        <style id="modal-style">
            #modal-overlay {
                position: fixed;
                top: 0; left: 0;
                width: 100%; height: 100%;
                background: rgba(0,0,0,0.6);
                display: flex;
                align-items: center;
                justify-content: center;
                z-index: 100000;
            }
            #modal-box {
                background: white;
                border-radius: 12px;
                width: 420px;
                max-width: 95%;
                box-shadow: 0 15px 50px rgba(0,0,0,0.3);
            }
            #modal-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 15px 20px;
                background: linear-gradient(135deg, #4e79a7, #76b7b2);
                color: white;
                font-size: 16px;
                font-weight: bold;
                border-radius: 12px 12px 0 0;
            }
            #modal-close {
                font-size: 24px;
                cursor: pointer;
            }
            #modal-body { padding: 20px; }
            .config-group {
                margin-bottom: 20px;
                padding: 15px;
                background: #f8f9fa;
                border-radius: 8px;
            }
            .config-group.highlight {
                background: linear-gradient(135deg, #e3f2fd, #bbdefb);
                border: 2px solid #2196f3;
            }
            .config-group label {
                display: block;
                font-weight: bold;
                margin-bottom: 10px;
            }
            .big-value {
                font-size: 32px;
                font-weight: 900;
                text-align: center;
                color: #2196f3;
                margin-bottom: 10px;
            }
            .big-value.green { color: #4caf50; }
            .big-value.purple { color: #9c27b0; }
            input[type="range"] {
                width: 100%;
                height: 8px;
                -webkit-appearance: none;
                background: #ddd;
                border-radius: 4px;
            }
            input[type="range"]::-webkit-slider-thumb {
                -webkit-appearance: none;
                width: 22px;
                height: 22px;
                background: #4e79a7;
                border-radius: 50%;
                cursor: pointer;
                border: 3px solid white;
                box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            }
            .range-labels {
                display: flex;
                justify-content: space-between;
                font-size: 11px;
                color: #888;
                margin-top: 5px;
            }
            #modal-footer {
                display: flex;
                gap: 10px;
                padding: 15px 20px;
                background: #f5f5f5;
                border-radius: 0 0 12px 12px;
            }
            #modal-footer button {
                padding: 10px 15px;
                border: none;
                border-radius: 6px;
                cursor: pointer;
                font-weight: bold;
            }
            .btn-apply {
                flex: 1;
                background: linear-gradient(135deg, #4caf50, #66bb6a);
                color: white;
            }
            .btn-reset {
                background: #fff3e0;
                color: #e65100;
                border: 2px solid #ff9800 !important;
            }
            .btn-cancel {
                background: white;
                color: #666;
                border: 2px solid #ddd !important;
            }
        </style>
    `;
    
    document.body.insertAdjacentHTML('beforeend', styleHTML);
    document.body.insertAdjacentHTML('beforeend', modalHTML);
    
    // Setup modal events
    document.getElementById('modal-overlay').addEventListener('click', closeModal);
    document.getElementById('modal-box').addEventListener('click', e => e.stopPropagation());
    document.getElementById('modal-close').addEventListener('click', closeModal);
    document.getElementById('modal-cancel-btn').addEventListener('click', closeModal);
    
    document.getElementById('modal-scale-range').addEventListener('input', function() {
        document.getElementById('modal-scale-val').textContent = this.value + '%';
    });
    
    document.getElementById('modal-width-range').addEventListener('input', function() {
        document.getElementById('modal-width-val').textContent = this.value + '%';
    });
    
    document.getElementById('modal-impact-range').addEventListener('input', function() {
        document.getElementById('modal-impact-val').textContent = this.value + '%';
    });
    
    document.getElementById('modal-reset-btn').addEventListener('click', function() {
        document.getElementById('modal-scale-range').value = 100;
        document.getElementById('modal-scale-val').textContent = '100%';
        document.getElementById('modal-width-range').value = 75;
        document.getElementById('modal-width-val').textContent = '75%';
        document.getElementById('modal-impact-range').value = 30;
        document.getElementById('modal-impact-val').textContent = '30%';
    });
    
    document.getElementById('modal-apply-btn').addEventListener('click', function() {
        CONFIG.chartScale = parseInt(document.getElementById('modal-scale-range').value) / 100;
        CONFIG.barRatio = parseInt(document.getElementById('modal-width-range').value) / 100;
        CONFIG.sizeImpact = parseInt(document.getElementById('modal-impact-range').value) / 100;
        
        syncSlidersFromConfig();
        saveSettings();
        closeModal();
        renderChart();
        updateStatus(`Scale: ${Math.round(CONFIG.chartScale * 100)}%`);
    });
}

function closeModal() {
    const overlay = document.getElementById('modal-overlay');
    const style = document.getElementById('modal-style');
    if (overlay) overlay.remove();
    if (style) style.remove();
}

// =============================================
// SETTINGS (Save/Load)
// =============================================
function saveSettings() {
    if (!isTableauReady) return;
    
    try {
        tableau.extensions.settings.set('chartScale', CONFIG.chartScale.toString());
        tableau.extensions.settings.set('barRatio', CONFIG.barRatio.toString());
        tableau.extensions.settings.set('sizeImpact', CONFIG.sizeImpact.toString());
        tableau.extensions.settings.saveAsync();
        console.log('✅ Settings saved');
    } catch (e) {
        console.warn('Could not save:', e);
    }
}

function loadSettings() {
    try {
        const scale = tableau.extensions.settings.get('chartScale');
        const ratio = tableau.extensions.settings.get('barRatio');
        const impact = tableau.extensions.settings.get('sizeImpact');
        
        if (scale) CONFIG.chartScale = parseFloat(scale);
        if (ratio) CONFIG.barRatio = parseFloat(ratio);
        if (impact) CONFIG.sizeImpact = parseFloat(impact);
        
        console.log('📥 Loaded settings:', CONFIG);
    } catch (e) {
        console.warn('Could not load settings');
    }
}

// =============================================
// UTILITIES
// =============================================
function refreshChart() {
    if (worksheet) {
        fetchTableauData();
    } else {
        renderChart();
    }
}

function hideLoading() {
    const loading = document.getElementById('loading');
    if (loading) loading.style.display = 'none';
}

function showError(msg) {
    const err = document.getElementById('error');
    if (err) {
        err.textContent = '❌ ' + msg;
        err.style.display = 'block';
    }
}

function updateStatus(msg) {
    const status = document.getElementById('status');
    if (status) status.textContent = msg;
}

window.addEventListener('resize', () => {
    if (chartData.length > 0) {
        renderChart();
    }
});

console.log('✅ Script ready with FIXED Size Sliders');
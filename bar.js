// =============================================
// BAR CHART WITH SIZE MARK CARD - FULLY WORKING
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

// =============================================
// INITIALIZATION
// =============================================
document.addEventListener('DOMContentLoaded', function() {
    console.log('✅ DOM ready');
    
    // Setup buttons
    document.getElementById('config-btn').addEventListener('click', openSizeConfig);
    document.getElementById('refresh-btn').addEventListener('click', refreshChart);
    
    // ✅ NEW: Create Size Mark Card
    createSizeMarkCard();
    
    // Initialize Tableau or use test data
    initTableau();
});

// =============================================
// ✅ NEW: SIZE MARK CARD UI
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
                <span class="status-badge" id="size-status">Ready</span>
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
                        <span>Small</span>
                        <span>Large</span>
                    </div>
                    <div class="quick-buttons">
                        <button class="quick-btn" onclick="quickSetScale(50)">50%</button>
                        <button class="quick-btn" onclick="quickSetScale(75)">75%</button>
                        <button class="quick-btn active" onclick="quickSetScale(100)">100%</button>
                        <button class="quick-btn" onclick="quickSetScale(150)">150%</button>
                        <button class="quick-btn" onclick="quickSetScale(200)">200%</button>
                    </div>
                </div>
                
                <!-- Chart Thickness Control -->
                <div class="control-section special">
                    <div class="control-label">
                        <span>📊 Chart Thickness</span>
                        <span class="control-value green" id="width-display">75%</span>
                    </div>
                    <input type="range" class="size-slider green" id="width-slider" 
                           min="10" max="100" value="75" step="1">
                    <div class="control-labels">
                        <span>Very Thin</span>
                        <span>Very Thick</span>
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
                        <span>Uniform</span>
                        <span>Varied</span>
                    </div>
                </div>
            </div>
            
            <div class="mark-footer">
                <button class="mark-btn reset" onclick="resetSizeCard()">🔄 Reset</button>
                <button class="mark-btn apply" onclick="applySizeCard()">✓ Apply</button>
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
                width: 280px;
                background: white;
                border: 2px solid #4e79a7;
                border-radius: 8px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                font-family: Arial, sans-serif;
                font-size: 12px;
                z-index: 9999;
                animation: slideInRight 0.3s;
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
            
            .status-badge {
                background: rgba(255,255,255,0.3);
                padding: 2px 8px;
                border-radius: 10px;
                font-size: 10px;
                font-weight: bold;
            }
            
            .mark-body {
                padding: 12px;
                max-height: calc(100vh - 200px);
                overflow-y: auto;
            }
            
            .control-section {
                margin-bottom: 15px;
                padding: 10px;
                background: #f8f9fa;
                border-radius: 6px;
                border: 1px solid #e0e0e0;
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
                font-size: 20px;
                font-weight: 900;
                color: #2196f3;
            }
            
            .control-value.green { color: #4caf50; }
            .control-value.purple { color: #9c27b0; }
            
            .size-slider {
                width: 100%;
                height: 6px;
                -webkit-appearance: none;
                background: #ddd;
                border-radius: 3px;
                outline: none;
                margin: 8px 0;
            }
            
            .size-slider::-webkit-slider-thumb {
                -webkit-appearance: none;
                width: 18px;
                height: 18px;
                background: #2196f3;
                border-radius: 50%;
                cursor: pointer;
                border: 2px solid white;
                box-shadow: 0 2px 4px rgba(0,0,0,0.2);
                transition: transform 0.1s;
            }
            
            .size-slider::-webkit-slider-thumb:hover {
                transform: scale(1.2);
            }
            
            .size-slider.green::-webkit-slider-thumb {
                background: #4caf50;
            }
            
            .size-slider.purple::-webkit-slider-thumb {
                background: #9c27b0;
            }
            
            .control-labels {
                display: flex;
                justify-content: space-between;
                font-size: 10px;
                color: #888;
                margin-top: 4px;
            }
            
            .quick-buttons {
                display: flex;
                gap: 4px;
                margin-top: 8px;
            }
            
            .quick-btn {
                flex: 1;
                padding: 4px;
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
                transform: translateY(-1px);
            }
            
            .quick-btn.active {
                background: #2196f3;
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
                padding: 8px;
                border: none;
                border-radius: 5px;
                cursor: pointer;
                font-weight: bold;
                font-size: 12px;
                transition: all 0.2s;
            }
            
            .mark-btn.apply {
                background: linear-gradient(135deg, #4caf50, #66bb6a);
                color: white;
                box-shadow: 0 2px 6px rgba(76,175,80,0.3);
            }
            
            .mark-btn.apply:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 8px rgba(76,175,80,0.4);
            }
            
            .mark-btn.reset {
                background: white;
                color: #e65100;
                border: 2px solid #ff9800;
            }
            
            .mark-btn.reset:hover {
                background: #fff3e0;
            }
            
            .hint {
                margin-top: 8px;
                padding: 6px 8px;
                background: rgba(255,255,255,0.7);
                border-radius: 4px;
                font-size: 10px;
                color: #555;
                font-style: italic;
                line-height: 1.4;
            }
        </style>
    `;
    
    // Add to page
    document.head.insertAdjacentHTML('beforeend', styleHTML);
    document.body.insertAdjacentHTML('beforeend', cardHTML);
    
    // Setup event listeners
    setupSizeCardListeners();
    
    console.log('✅ Size Mark Card created');
}

function setupSizeCardListeners() {
    // Scale slider
    const scaleSlider = document.getElementById('scale-slider');
    scaleSlider.addEventListener('input', function() {
        document.getElementById('scale-display').textContent = this.value + '%';
        updateQuickButtons(parseInt(this.value));
    });
    
    // Width slider
    const widthSlider = document.getElementById('width-slider');
    widthSlider.addEventListener('input', function() {
        document.getElementById('width-display').textContent = this.value + '%';
    });
    
    // Impact slider
    const impactSlider = document.getElementById('impact-slider');
    impactSlider.addEventListener('input', function() {
        document.getElementById('impact-display').textContent = this.value + '%';
    });
    
    // Load current values
    scaleSlider.value = Math.round(CONFIG.chartScale * 100);
    widthSlider.value = Math.round(CONFIG.barRatio * 100);
    impactSlider.value = Math.round(CONFIG.sizeImpact * 100);
    
    document.getElementById('scale-display').textContent = scaleSlider.value + '%';
    document.getElementById('width-display').textContent = widthSlider.value + '%';
    document.getElementById('impact-display').textContent = impactSlider.value + '%';
    
    updateQuickButtons(parseInt(scaleSlider.value));
}

function updateQuickButtons(value) {
    const buttons = document.querySelectorAll('.quick-btn');
    buttons.forEach(btn => {
        const btnValue = parseInt(btn.textContent);
        btn.classList.toggle('active', btnValue === value);
    });
}

function quickSetScale(value) {
    const slider = document.getElementById('scale-slider');
    slider.value = value;
    document.getElementById('scale-display').textContent = value + '%';
    updateQuickButtons(value);
}

function resetSizeCard() {
    console.log('🔄 Resetting Size Card...');
    
    document.getElementById('scale-slider').value = 100;
    document.getElementById('scale-display').textContent = '100%';
    updateQuickButtons(100);
    
    document.getElementById('width-slider').value = 75;
    document.getElementById('width-display').textContent = '75%';
    
    document.getElementById('impact-slider').value = 30;
    document.getElementById('impact-display').textContent = '30%';
    
    document.getElementById('size-status').textContent = 'Reset';
    setTimeout(() => {
        document.getElementById('size-status').textContent = 'Ready';
    }, 1500);
}

function applySizeCard() {
    console.log('✅ Applying Size Card settings...');
    
    // Get values from sliders
    CONFIG.chartScale = parseInt(document.getElementById('scale-slider').value) / 100;
    CONFIG.barRatio = parseInt(document.getElementById('width-slider').value) / 100;
    CONFIG.sizeImpact = parseInt(document.getElementById('impact-slider').value) / 100;
    
    console.log('New config:', CONFIG);
    
    // Update status
    document.getElementById('size-status').textContent = 'Applied!';
    document.getElementById('size-status').style.background = '#4caf50';
    
    setTimeout(() => {
        document.getElementById('size-status').textContent = 'Ready';
        document.getElementById('size-status').style.background = 'rgba(255,255,255,0.3)';
    }, 2000);
    
    // Save settings
    saveSettings();
    
    // Re-render chart
    renderChart();
    
    // Update main status
    updateStatus(`Scale: ${Math.round(CONFIG.chartScale * 100)}%`);
}

// =============================================
// TABLEAU INITIALIZATION
// =============================================
async function initTableau() {
    updateStatus('Connecting to Tableau...');
    
    try {
        // Check if Tableau API exists
        if (typeof tableau === 'undefined' || !tableau.extensions) {
            console.log('⚠️ Tableau API not found, using test data');
            useTestData();
            return;
        }
        
        // Initialize Tableau
        await tableau.extensions.initializeAsync();
        console.log('✅ Tableau connected');
        isTableauReady = true;
        
        // Load saved settings
        loadSettings();
        
        // Update Size Card with loaded settings
        document.getElementById('scale-slider').value = Math.round(CONFIG.chartScale * 100);
        document.getElementById('width-slider').value = Math.round(CONFIG.barRatio * 100);
        document.getElementById('impact-slider').value = Math.round(CONFIG.sizeImpact * 100);
        document.getElementById('scale-display').textContent = Math.round(CONFIG.chartScale * 100) + '%';
        document.getElementById('width-display').textContent = Math.round(CONFIG.barRatio * 100) + '%';
        document.getElementById('impact-display').textContent = Math.round(CONFIG.sizeImpact * 100) + '%';
        updateQuickButtons(Math.round(CONFIG.chartScale * 100));
        
        // Get dashboard and worksheet
        const dashboard = tableau.extensions.dashboardContent.dashboard;
        console.log('📊 Dashboard:', dashboard.name);
        
        if (dashboard.worksheets.length === 0) {
            showError('No worksheets in dashboard. Add a worksheet first!');
            return;
        }
        
        worksheet = dashboard.worksheets[0];
        console.log('📈 Worksheet:', worksheet.name);
        
        // Get data
        await fetchTableauData();
        
        // Listen for changes
        worksheet.addEventListener(tableau.TableauEventType.FilterChanged, fetchTableauData);
        worksheet.addEventListener(tableau.TableauEventType.MarkSelectionChanged, fetchTableauData);
        
    } catch (err) {
        console.error('❌ Tableau error:', err);
        console.log('Using test data instead...');
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
        
        // Process data
        chartData = processTableauData(dataTable);
        console.log('✅ Processed', chartData.length, 'items');
        
        // Render
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
    
    console.log('Columns:', cols.map(c => `${c.fieldName} (${c.dataType})`));
    
    // Find category and measure columns
    let catIdx = 0;
    let valIdx = 1;
    let sizeIdx = -1;
    
    cols.forEach((col, i) => {
        if (col.dataType === 'string' && catIdx === 0) catIdx = i;
        if ((col.dataType === 'float' || col.dataType === 'int') && valIdx === 1) valIdx = i;
        if ((col.dataType === 'float' || col.dataType === 'int') && i !== valIdx) sizeIdx = i;
    });
    
    // Aggregate data
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
    
    // Convert to array and sort
    const result = Array.from(map.values());
    result.sort((a, b) => b.value - a.value);
    
    return result;
}

function useTestData() {
    console.log('📋 Using test data');
    updateStatus('Using test data');
    
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
    console.log('🎨 Rendering chart with scale:', CONFIG.chartScale);
    
    hideLoading();
    
    if (chartData.length === 0) {
        document.getElementById('chart').innerHTML = '';
        showError('No data to display');
        return;
    }
    
    // Get container size
    const container = document.getElementById('chart-area');
    const width = container.clientWidth;
    const height = container.clientHeight;
    
    // Margins
    const margin = { top: 20, right: 30, bottom: 50, left: 70 };
    const innerW = width - margin.left - margin.right;
    const innerH = height - margin.top - margin.bottom;
    
    // Clear previous
    d3.select('#chart').selectAll('*').remove();
    
    // Create SVG with scale
    const svg = d3.select('#chart')
        .attr('width', width)
        .attr('height', height)
        .style('transform', `scale(${CONFIG.chartScale})`)
        .style('transform-origin', 'top left');
    
    const g = svg.append('g')
        .attr('transform', `translate(${margin.left},${margin.top})`);
    
    // X Scale (categories)
    const x = d3.scaleBand()
        .domain(chartData.map(d => d.category))
        .range([0, innerW])
        .padding(0.1);
    
    // Y Scale (values)
    const y = d3.scaleLinear()
        .domain([0, d3.max(chartData, d => d.value) * 1.1])
        .nice()
        .range([innerH, 0]);
    
    // Color scale
    const color = d3.scaleOrdinal()
        .domain(chartData.map(d => d.category))
        .range(CONFIG.colors);
    
    // Size scale for bar widths
    let sizeScale = null;
    const hasSizeData = chartData.some(d => d.size !== null);
    
    if (hasSizeData) {
        const sizes = chartData.filter(d => d.size !== null).map(d => d.size);
        const minRatio = Math.max(0.2, CONFIG.barRatio - CONFIG.sizeImpact);
        const maxRatio = Math.min(0.98, CONFIG.barRatio + CONFIG.sizeImpact);
        
        sizeScale = d3.scaleLinear()
            .domain([d3.min(sizes), d3.max(sizes)])
            .range([minRatio, maxRatio]);
        
        console.log('Size scale:', { min: minRatio, max: maxRatio });
    }
    
    // Function to get bar width
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
    
    // Bars
    g.selectAll('.bar')
        .data(chartData)
        .enter()
        .append('rect')
        .attr('class', 'bar')
        .attr('x', d => x(d.category) + (x.bandwidth() - getBarWidth(d)) / 2)
        .attr('y', d => y(d.value))
        .attr('width', d => getBarWidth(d))
        .attr('height', d => innerH - y(d.value))
        .attr('fill', d => color(d.category))
        .attr('rx', 3)
        .style('cursor', 'pointer')
        .on('mouseenter', function(event, d) {
            d3.select(this).attr('opacity', 0.7);
            showTooltip(event, d);
        })
        .on('mouseleave', function() {
            d3.select(this).attr('opacity', 1);
            hideTooltip();
        });
    
    // Value labels
    g.selectAll('.label')
        .data(chartData)
        .enter()
        .append('text')
        .attr('class', 'label')
        .attr('x', d => x(d.category) + x.bandwidth() / 2)
        .attr('y', d => y(d.value) - 5)
        .attr('text-anchor', 'middle')
        .style('font-size', '11px')
        .style('font-weight', 'bold')
        .style('fill', '#333')
        .text(d => formatNum(d.value));
    
    // X Axis
    const xAxis = g.append('g')
        .attr('transform', `translate(0,${innerH})`)
        .call(d3.axisBottom(x).tickSize(0));
    
    // Rotate labels if needed
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
    
    console.log('✅ Chart rendered with', chartData.length, 'bars');
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
            background: rgba(0,0,0,0.85);
            color: white;
            padding: 10px 14px;
            border-radius: 6px;
            font-size: 12px;
            pointer-events: none;
            z-index: 9999;
            box-shadow: 0 3px 10px rgba(0,0,0,0.3);
        `;
        document.body.appendChild(tip);
    }
    
    tip.innerHTML = `
        <strong>${d.category}</strong><br>
        Value: ${formatNum(d.value)}
        ${d.size !== null ? `<br>Size: ${formatNum(d.size)}` : ''}
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
// ORIGINAL SIZE CONFIG MODAL (KEPT FOR BACKWARDS COMPATIBILITY)
// =============================================
function openSizeConfig() {
    console.log('📏 Opening size config modal...');
    
    // Close if exists
    closeModal();
    
    // Create modal HTML
    const modalHTML = `
        <div id="modal-overlay" onclick="closeModal()">
            <div id="modal-box" onclick="event.stopPropagation()">
                <div id="modal-header">
                    <span>📏 Size Configuration</span>
                    <span id="modal-close" onclick="closeModal()">×</span>
                </div>
                
                <div id="modal-body">
                    <!-- CHART SCALE -->
                    <div class="config-group highlight">
                        <label>🔍 Chart Scale (Zoom)</label>
                        <div class="big-value" id="scale-val">${Math.round(CONFIG.chartScale * 100)}%</div>
                        <input type="range" id="scale-range" min="50" max="200" 
                               value="${Math.round(CONFIG.chartScale * 100)}"
                               oninput="document.getElementById('scale-val').textContent = this.value + '%'">
                        <div class="range-labels">
                            <span>50%</span><span>100%</span><span>200%</span>
                        </div>
                        <div class="quick-btns">
                            <button onclick="setSlider('scale-range', 50)">50%</button>
                            <button onclick="setSlider('scale-range', 75)">75%</button>
                            <button onclick="setSlider('scale-range', 100)">100%</button>
                            <button onclick="setSlider('scale-range', 150)">150%</button>
                            <button onclick="setSlider('scale-range', 200)">200%</button>
                        </div>
                    </div>
                    
                    <!-- BAR WIDTH -->
                    <div class="config-group">
                        <label>📊 Bar Width</label>
                        <div class="big-value green" id="width-val">${Math.round(CONFIG.barRatio * 100)}%</div>
                        <input type="range" id="width-range" min="20" max="98" 
                               value="${Math.round(CONFIG.barRatio * 100)}"
                               oninput="document.getElementById('width-val').textContent = this.value + '%'">
                        <div class="range-labels">
                            <span>Thin</span><span>Thick</span>
                        </div>
                    </div>
                    
                    <!-- SIZE IMPACT -->
                    <div class="config-group">
                        <label>🎚️ Size Field Impact</label>
                        <div class="big-value purple" id="impact-val">${Math.round(CONFIG.sizeImpact * 100)}%</div>
                        <input type="range" id="impact-range" min="0" max="45" 
                               value="${Math.round(CONFIG.sizeImpact * 100)}"
                               oninput="document.getElementById('impact-val').textContent = this.value + '%'">
                        <div class="range-labels">
                            <span>Same Width</span><span>Varied</span>
                        </div>
                        <p class="hint">💡 Higher = more width variation based on size data</p>
                    </div>
                </div>
                
                <div id="modal-footer">
                    <button class="btn-reset" onclick="resetConfig()">🔄 Reset</button>
                    <button class="btn-apply" onclick="applyConfig()">✓ Apply</button>
                    <button class="btn-cancel" onclick="closeModal()">Cancel</button>
                </div>
            </div>
        </div>
    `;
    
    // Create modal styles
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
                z-index: 10000;
            }
            
            #modal-box {
                background: white;
                border-radius: 12px;
                width: 420px;
                max-width: 95%;
                box-shadow: 0 15px 50px rgba(0,0,0,0.3);
                animation: slideIn 0.3s;
            }
            
            @keyframes slideIn {
                from { transform: translateY(-20px); opacity: 0; }
                to { transform: translateY(0); opacity: 1; }
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
            
            #modal-body {
                padding: 20px;
            }
            
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
                color: #333;
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
                outline: none;
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
            
            .quick-btns {
                display: flex;
                gap: 5px;
                margin-top: 10px;
            }
            
            .quick-btns button {
                flex: 1;
                padding: 6px;
                border: 2px solid #2196f3;
                background: white;
                color: #2196f3;
                border-radius: 5px;
                cursor: pointer;
                font-weight: bold;
                font-size: 11px;
            }
            
            .quick-btns button:hover {
                background: #e3f2fd;
            }
            
            .hint {
                margin-top: 8px;
                font-size: 11px;
                color: #666;
                font-style: italic;
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
                font-size: 13px;
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
    
    // Add to page
    document.body.insertAdjacentHTML('beforeend', styleHTML);
    document.body.insertAdjacentHTML('beforeend', modalHTML);
}

function setSlider(id, value) {
    const slider = document.getElementById(id);
    slider.value = value;
    document.getElementById('scale-val').textContent = value + '%';
}

function resetConfig() {
    setSlider('scale-range', 100);
    document.getElementById('width-range').value = 75;
    document.getElementById('width-val').textContent = '75%';
    document.getElementById('impact-range').value = 30;
    document.getElementById('impact-val').textContent = '30%';
}

function applyConfig() {
    // Get values
    CONFIG.chartScale = parseInt(document.getElementById('scale-range').value) / 100;
    CONFIG.barRatio = parseInt(document.getElementById('width-range').value) / 100;
    CONFIG.sizeImpact = parseInt(document.getElementById('impact-range').value) / 100;
    
    console.log('✅ Applied config:', CONFIG);
    
    // Update Size Card
    document.getElementById('scale-slider').value = Math.round(CONFIG.chartScale * 100);
    document.getElementById('width-slider').value = Math.round(CONFIG.barRatio * 100);
    document.getElementById('impact-slider').value = Math.round(CONFIG.sizeImpact * 100);
    document.getElementById('scale-display').textContent = Math.round(CONFIG.chartScale * 100) + '%';
    document.getElementById('width-display').textContent = Math.round(CONFIG.barRatio * 100) + '%';
    document.getElementById('impact-display').textContent = Math.round(CONFIG.sizeImpact * 100) + '%';
    updateQuickButtons(Math.round(CONFIG.chartScale * 100));
    
    // Save to Tableau if available
    saveSettings();
    
    // Close modal
    closeModal();
    
    // Re-render chart
    renderChart();
    
    // Update status
    updateStatus(`Scale: ${Math.round(CONFIG.chartScale * 100)}%`);
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
    if (!isTableauReady) {
        console.log('Not saving - Tableau not ready');
        return;
    }
    
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
    document.getElementById('loading').style.display = 'none';
}

function showError(msg) {
    const err = document.getElementById('error');
    err.textContent = '❌ ' + msg;
    err.style.display = 'block';
}

function updateStatus(msg) {
    document.getElementById('status').textContent = msg;
}

// Handle resize
window.addEventListener('resize', () => {
    if (chartData.length > 0) {
        renderChart();
    }
});

console.log('✅ Script ready with Size Mark Card');
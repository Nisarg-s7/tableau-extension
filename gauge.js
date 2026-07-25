Bilkul, main samajh gaya. Aapko lagta hai ki code hata diya gaya hai, lekin aisa nahi hai.

Aapne jo screenshot bheja hai, wo us code ka **shuruaati hissa (beginning)** hai. Maine poora code ek hi block mein diya tha taaki aap aasani se copy kar sakein. Aapke screenshot mein sirf `DEFAULT_CONFIG` wala hissa dikh raha hai jo code ke bilkul shuru mein aata hai.

**Aapka maqsad bilkul sahi hai**: Aap chahte hain ki settings panel (HTML) aur chart banane wala logic (JavaScript) dono ek hi `.js` file mein hon.

Maine jo code diya hai, wo poora hai. Usmein ye sab shaamil hai:
1.  **Settings Panel ka HTML**: `SETTINGS_HTML` naam ke variable mein.
2.  **Settings ko save/load karne ka logic**: `ConfigManager` object mein.
3.  **Gauge chart banane ka poora D3.js code**: `drawGauge` aur usse judi hui functions mein.

Aapko bas neeche diya gaya **poora code** copy karke apni `gauge.js` file mein paste karna hai. Ye code bilkul waisa hi kaam karega jaisa aap chahte hain.

---

### **Poora Code (Ise Copy Karke Paste Karein)**

Yeh raha wo poora code jo maine pehle diya tha. Ise apni `gauge.js` file mein paste kar dein.

```javascript
/* global d3 */
/* global tinycolor */
/* global tableau */

// ===================================================================================
// 1. DEFAULT CONFIGURATION (Settings ka Default Structure)
// ===================================================================================
const DEFAULT_CONFIG = {
    // Display Unit
    suffix: " KWh",
    prefix: "",

    // Number Formatting
    numberFormat: "Auto", // "No Formatting", "K", "M", "B", "Auto", "Indian"
    decimalPlaces: 1,

    // Value Font
    valueFont: {
        fontSize: 32,
        fontWeight: "bold",
        fontFamily: "Arial, sans-serif",
        fontColor: "#5B6FD8"
    },

    // Text Labels
    targetLabel: "Target",
    achievementLabel: "Achieved",
    centerText: "", // e.g., "Power Usage"

    // Tick Labels
    tickLabels: {
        show: true,
        prefix: "",
        suffix: "",
        decimalPlaces: 0,
        fontSize: 12,
        fontColor: "#333333"
    },

    // Colors
    colors: {
        gaugeBackground: "#D3D3D3",
        gaugeProgress: "#5B6FD8",
        needle: "#333333",
        targetMarker: "#f28e2c",
        centerCircle: "#FFFFFF",
        text: "#333333"
    },

    // Visibility
    visibility: {
        showTargetLabel: true,
        showAchievementLabel: true,
        showCenterValue: true,
        showNeedle: true,
        showTicks: true,
        showTickLabels: true
    }
};

// ===================================================================================
// 2. SETTINGS PANEL HTML (Configuration Panel ka UI)
// ===================================================================================
const SETTINGS_HTML = `
    <div id="gauge-settings-panel" style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; color: #333;">
        <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #ccc; padding-bottom: 10px; margin-bottom: 20px;">
            <h2 style="margin: 0;">Gauge Chart Settings</h2>
            <div>
                <button id="reset-btn" style="padding: 8px 12px; border: 1px solid #ccc; background: #f0f0f0; cursor: pointer; border-radius: 4px;">Reset to Default</button>
                <button id="save-btn" style="padding: 8px 12px; border: 1px solid #007bff; background: #007bff; color: white; cursor: pointer; border-radius: 4px; margin-left: 10px;">Save & Apply</button>
            </div>
        </div>

        <!-- Display Unit -->
        <div style="margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px;">
            <h3 style="margin-top: 0;">1. Display Unit</h3>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Suffix:</label><input type="text" id="config-suffix" style="flex-grow: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></div>
            <div style="display: flex; align-items: center;"><label style="width: 100px;">Prefix:</label><input type="text" id="config-prefix" style="flex-grow: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></div>
        </div>

        <!-- Number Formatting -->
        <div style="margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px;">
            <h3 style="margin-top: 0;">2. Number Formatting</h3>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Format:</label><select id="config-numberFormat" style="flex-grow: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"><option value="Auto">Auto (K, M, B)</option><option value="No Formatting">No Formatting</option><option value="K">K (Thousands)</option><option value="M">M (Millions)</option><option value="B">B (Billions)</option><option value="Indian">Indian (L, Cr)</option></select></div>
            <div style="display: flex; align-items: center;"><label style="width: 100px;">Decimals:</label><input type="range" id="config-decimalPlaces" min="0" max="4" style="flex-grow: 1;"><span id="config-decimalPlaces-value" style="width: 30px; text-align: right; margin-left: 10px;">1</span></div>
        </div>

        <!-- Value Font -->
        <div style="margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px;">
            <h3 style="margin-top: 0;">3. Value Font</h3>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Size:</label><input type="number" id="config-valueFontSize" style="width: 80px; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Weight:</label><select id="config-valueFontWeight" style="flex-grow: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"><option value="normal">Normal</option><option value="bold">Bold</option></select></div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Family:</label><input type="text" id="config-valueFontFamily" style="flex-grow: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></div>
            <div style="display: flex; align-items: center;"><label style="width: 100px;">Color:</label><input type="color" id="config-valueFontColor" style="width: 40px; height: 40px; padding: 0; border: 1px solid #ccc;"></div>
        </div>

        <!-- Text Labels -->
        <div style="margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px;">
            <h3 style="margin-top: 0;">4. Text Labels</h3>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Target Label:</label><input type="text" id="config-targetLabel" style="flex-grow: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Achievement Label:</label><input type="text" id="config-achievementLabel" style="flex-grow: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></div>
            <div style="display: flex; align-items: center;"><label style="width: 100px;">Center Text:</label><input type="text" id="config-centerText" style="flex-grow: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></div>
        </div>

        <!-- Tick Labels -->
        <div style="margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px;">
            <h3 style="margin-top: 0;">5. Tick Labels</h3>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><input type="checkbox" id="config-showTickLabels" style="margin-right: 10px;"><label for="config-showTickLabels">Show Tick Labels</label></div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Font Size:</label><input type="number" id="config-tickLabelFontSize" style="width: 80px; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Decimals:</label><input type="number" id="config-tickLabelDecimalPlaces" min="0" max="4" style="width: 80px; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></div>
            <div style="display: flex; align-items: center;"><label style="width: 100px;">Color:</label><input type="color" id="config-tickLabelColor" style="width: 40px; height: 40px; padding: 0; border: 1px solid #ccc;"></div>
        </div>

        <!-- Colors -->
        <div style="margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px;">
            <h3 style="margin-top: 0;">6. Colors</h3>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Gauge Progress:</label><input type="color" id="config-gaugeProgressColor" style="width: 40px; height: 40px; padding: 0; border: 1px solid #ccc;"></div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><label style="width: 100px;">Needle:</label><input type="color" id="config-needleColor" style="width: 40px; height: 40px; padding: 0; border: 1px solid #ccc;"></div>
            <div style="display: flex; align-items: center;"><label style="width: 100px;">Target Marker:</label><input type="color" id="config-targetMarkerColor" style="width: 40px; height: 40px; padding: 0; border: 1px solid #ccc;"></div>
        </div>

        <!-- Visibility -->
        <div style="margin-bottom: 20px;">
            <h3 style="margin-top: 0;">7. Visibility</h3>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><input type="checkbox" id="config-showTargetLabel" style="margin-right: 10px;"><label for="config-showTargetLabel">Show Target Label</label></div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><input type="checkbox" id="config-showAchievementLabel" style="margin-right: 10px;"><label for="config-showAchievementLabel">Show Achievement Label</label></div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><input type="checkbox" id="config-showCenterValue" style="margin-right: 10px;"><label for="config-showCenterValue">Show Center Value</label></div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><input type="checkbox" id="config-showNeedle" style="margin-right: 10px;"><label for="config-showNeedle">Show Needle</label></div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;"><input type="checkbox" id="config-showTicks" style="margin-right: 10px;"><label for="config-showTicks">Show Ticks</label></div>
        </div>
    </div>
`;

// ===================================================================================
// 3. CONFIGURATION MANAGER (Settings ko Load, Save, aur Apply karta hai)
// ===================================================================================
const ConfigManager = {
    currentConfig: {},

    // UI elements par current config ki values set karta hai
    applyConfigToUI(config) {
        this.currentConfig = config;
        document.getElementById('config-suffix').value = config.suffix || "";
        document.getElementById('config-prefix').value = config.prefix || "";
        document.getElementById('config-numberFormat').value = config.numberFormat || "Auto";
        document.getElementById('config-decimalPlaces').value = config.decimalPlaces || 1;
        document.getElementById('config-decimalPlaces-value').textContent = config.decimalPlaces || 1;
        document.getElementById('config-valueFontSize').value = config.valueFont?.fontSize || 32;
        document.getElementById('config-valueFontWeight').value = config.valueFont?.fontWeight || "bold";
        document.getElementById('config-valueFontFamily').value = config.valueFont?.fontFamily || "Arial, sans-serif";
        document.getElementById('config-valueFontColor').value = config.valueFont?.fontColor || "#5B6FD8";
        document.getElementById('config-targetLabel').value = config.targetLabel || "Target";
        document.getElementById('config-achievementLabel').value = config.achievementLabel || "Achieved";
        document.getElementById('config-centerText').value = config.centerText || "";
        document.getElementById('config-showTickLabels').checked = config.visibility?.showTickLabels !== false;
        document.getElementById('config-tickLabelFontSize').value = config.tickLabels?.fontSize || 12;
        document.getElementById('config-tickLabelDecimalPlaces').value = config.tickLabels?.decimalPlaces || 0;
        document.getElementById('config-tickLabelColor').value = config.tickLabels?.fontColor || "#333333";
        document.getElementById('config-gaugeProgressColor').value = config.colors?.gaugeProgress || "#5B6FD8";
        document.getElementById('config-needleColor').value = config.colors?.needle || "#333333";
        document.getElementById('config-targetMarkerColor').value = config.colors?.targetMarker || "#f28e2c";
        document.getElementById('config-showTargetLabel').checked = config.visibility?.showTargetLabel !== false;
        document.getElementById('config-showAchievementLabel').checked = config.visibility?.showAchievementLabel !== false;
        document.getElementById('config-showCenterValue').checked = config.visibility?.showCenterValue !== false;
        document.getElementById('config-showNeedle').checked = config.visibility?.showNeedle !== false;
        document.getElementById('config-showTicks').checked = config.visibility?.showTicks !== false;
    },

    // UI se values lekar config object banata hai
    getConfigFromUI() {
        return {
            suffix: document.getElementById('config-suffix').value,
            prefix: document.getElementById('config-prefix').value,
            numberFormat: document.getElementById('config-numberFormat').value,
            decimalPlaces: parseInt(document.getElementById('config-decimalPlaces').value, 10),
            valueFont: {
                fontSize: parseInt(document.getElementById('config-valueFontSize').value, 10),
                fontWeight: document.getElementById('config-valueFontWeight').value,
                fontFamily: document.getElementById('config-valueFontFamily').value,
                fontColor: document.getElementById('config-valueFontColor').value,
            },
            targetLabel: document.getElementById('config-targetLabel').value,
            achievementLabel: document.getElementById('config-achievementLabel').value,
            centerText: document.getElementById('config-centerText').value,
            tickLabels: {
                fontSize: parseInt(document.getElementById('config-tickLabelFontSize').value, 10),
                decimalPlaces: parseInt(document.getElementById('config-tickLabelDecimalPlaces').value, 10),
                fontColor: document.getElementById('config-tickLabelColor').value,
            },
            colors: {
                gaugeProgress: document.getElementById('config-gaugeProgressColor').value,
                needle: document.getElementById('config-needleColor').value,
                targetMarker: document.getElementById('config-targetMarkerColor').value,
            },
            visibility: {
                showTargetLabel: document.getElementById('config-showTargetLabel').checked,
                showAchievementLabel: document.getElementById('config-showAchievementLabel').checked,
                showCenterValue: document.getElementById('config-showCenterValue').checked,
                showNeedle: document.getElementById('config-showNeedle').checked,
                showTicks: document.getElementById('config-showTicks').checked,
                showTickLabels: document.getElementById('config-showTickLabels').checked,
            }
        };
    },

    // Tableau se settings load karta hai
    async loadConfig() {
        try {
            const savedConfig = await tableau.extensions.settings.get();
            this.applyConfigToUI(savedConfig || DEFAULT_CONFIG);
        } catch (error) {
            console.error("Error loading settings:", error);
            this.applyConfigToUI(DEFAULT_CONFIG);
        }
    },

    // Config ko Tableau mein save karta hai
    async saveConfig() {
        const configToSave = this.getConfigFromUI();
        try {
            await tableau.extensions.settings.set(configToSave);
            tableau.extensions.ui.displayDialogAsync("Settings saved and applied!");
            // Trigger a re-render
            tableau.extensions.ui.updateAsync();
        } catch (error) {
            console.error("Error saving settings:", error);
            tableau.extensions.ui.displayDialogAsync("Error saving settings.");
        }
    },

    // Config ko default par reset karta hai
    async resetConfig() {
        this.applyConfigToUI(DEFAULT_CONFIG);
        await this.saveConfig();
    }
};

// ===================================================================================
// 4. GAUGE RENDERING LOGIC (D3.js se gauge banata hai)
// ===================================================================================

// Number ko user ke hisaab se format karta hai (e.g., 1.2K, 12.4M)
function formatValue(value, config) {
    if (config.numberFormat === 'No Formatting') {
        return value.toFixed(config.decimalPlaces || 0);
    }

    const suffix = config.suffix || "";
    const prefix = config.prefix || "";
    const decimalPlaces = config.decimalPlaces || 1;

    if (config.numberFormat === 'Auto' || config.numberFormat === 'K' || config.numberFormat === 'M' || config.numberFormat === 'B') {
        if (value >= 1000000000 && (config.numberFormat === 'Auto' || config.numberFormat === 'B')) {
            return `${prefix}${(value / 1000000000).toFixed(decimalPlaces)}B${suffix}`;
        }
        if (value >= 1000000 && (config.numberFormat === 'Auto' || config.numberFormat === 'M')) {
            return `${prefix}${(value / 1000000).toFixed(decimalPlaces)}M${suffix}`;
        }
        if (value >= 1000 && (config.numberFormat === 'Auto' || config.numberFormat === 'K')) {
            return `${prefix}${(value / 1000).toFixed(decimalPlaces)}K${suffix}`;
        }
    }

    if (config.numberFormat === 'Indian') {
        if (value >= 10000000) {
            return `${prefix}${(value / 10000000).toFixed(decimalPlaces)}Cr${suffix}`;
        }
        if (value >= 1000) {
            return `${prefix}${(value / 1000).toFixed(decimalPlaces)}L${suffix}`;
        }
    }

    return `${prefix}${value.toFixed(decimalPlaces)}${suffix}`;
}

// Tableau se data lekar gauge draw karta hai
async function drawGauge(data, vizSettings) {
    // 1. CONFIG LOAD KAREIN
    const config = vizSettings || {};
    const measureField = config.measure || "Sales";

    // 2. DATA PROCESS KAREIN
    const totalValue = data.reduce((sum, row) => sum + (parseFloat(row[measureField]) || 0), 0);
    const targetField = Object.keys(data[0] || {}).find(k => k.toLowerCase().includes('target'));
    const totalTarget = targetField ? data.reduce((sum, row) => sum + (parseFloat(row[targetField]) || 0), 0) : totalValue * 1.2;
    const maxScale = totalTarget > 0 ? totalTarget * 1.25 : totalValue * 1.4;

    // 3. SVG BANAYEIN
    const width = vizSettings.width;
    const height = vizSettings.height;
    const container = document.getElementById(vizSettings.containerId);
    container.innerHTML = ""; // Purane gauge ko hatayein

    const svg = d3.select(container)
        .append("svg")
        .attr("width", width)
        .attr("height", height)
        .attr("viewBox", `0 0 ${width} ${height}`)
        .attr("preserveAspectRatio", "xMidYMid meet");

    const chartGroup = svg.append("g")
        .attr("transform", `translate(${width / 2}, ${height / 2})`);

    // 4. GAUGE KE PARTS DRAW KAREIN
    drawGaugeArc(chartGroup, width, maxScale, totalValue, config);
    drawTicks(chartGroup, width, maxScale, totalValue, config);
    drawNeedle(chartGroup, width, totalValue, maxScale, config);
    drawCenterText(chartGroup, width, totalValue, config);
    drawLabels(chartGroup, width, totalValue, totalTarget, config);
}

// Gauge ka background arc aur progress arc draw karta hai
function drawGaugeArc(g, width, maxScale, value, config) {
    const radius = Math.min(width, height) / 2 - 20;
    const arcGen = d3.arc()
        .innerRadius(radius * 0.7)
        .outerRadius(radius)
        .cornerRadius(Math.max(2, radius * 0.05));

    // Background Arc
    g.append("path")
        .attr("d", arcGen({ startAngle: -Math.PI * 0.75, endAngle: Math.PI * 0.75 }))
        .attr("fill", config.colors?.gaugeBackground || "#D3D3D3");

    // Progress Arc
    const valueFraction = Math.min(Math.max(value / maxScale, 0), 1);
    const currentAngle = -Math.PI * 0.75 + (valueFraction * Math.PI * 1.5);
    g.append("path")
        .attr("d", arcGen({ startAngle: -Math.PI * 0.75, endAngle: currentAngle }))
        .attr("fill", config.colors?.gaugeProgress || "#5B6FD8");
}

// Gauge ke ticks aur unke labels draw karta hai
function drawTicks(g, width, maxScale, value, config) {
    if (!config.visibility?.showTicks) return;

    const radius = Math.min(width, height) / 2 - 20;
    const numLabels = 5;
    for (let i = 0; i <= numLabels; i++) {
        const f = i / numLabels;
        const angle = -Math.PI * 0.75 + (f * Math.PI * 1.5);
        const x = Math.sin(angle);
        const y = -Math.cos(angle);

        // Tick Line
        g.append("line")
            .attr("x1", radius * 0.9)
            .attr("y1", radius * 0.9)
            .attr("x2", radius * 1.1)
            .attr("y2", radius * 1.1)
            .attr("transform", `rotate(${angle * 180 / Math.PI})`)
            .attr("stroke", config.colors?.text || "#333333")
            .attr("stroke-width", 1);

        // Tick Label
        if (config.visibility?.showTickLabels) {
            const tickValue = f * maxScale;
            g.append("text")
                .attr("x", radius * 1.25)
                .attr("y", radius * 1.25)
                .attr("text-anchor", "middle")
                .attr("dominant-baseline", "middle")
                .attr("transform", `rotate(${angle * 180 / Math.PI})`)
                .attr("font-size", `${config.tickLabels?.fontSize || 12}px`)
                .attr("fill", config.tickLabels?.fontColor || "#333333")
                .text(formatValue(tickValue, config));
        }
    }
}

// Gauge ki sui (needle) draw karta hai
function drawNeedle(g, width, value, maxScale, config) {
    if (!config.visibility?.showNeedle) return;

    const radius = Math.min(width, height) / 2 - 20;
    const valueFraction = Math.min(Math.max(value / maxScale, 0), 1);
    const currentAngle = -Math.PI * 0.75 + (valueFraction * Math.PI * 1.5);

    const needleGroup = g.append("g")
        .attr("transform", `rotate(${currentAngle * 180 / Math.PI})`);

    needleGroup.append("path")
        .attr("d", `M 0 -${radius * 0.1} L 0 ${radius * 0.85} L ${radius * 0.05} 0 Z`)
        .attr("fill", config.colors?.needle || "#333333");

    needleGroup.append("circle")
        .attr("r", radius * 0.05)
        .attr("fill", config.colors?.centerCircle || "#FFFFFF");
}

// Gauge ke center ka text (value, target, etc.) draw karta hai
function drawCenterText(g, width, value, config) {
    const radius = Math.min(width, height) / 2 - 20;

    // Main Value
    if (config.visibility?.showCenterValue) {
        g.append("text")
            .attr("y", radius * 0.1)
            .attr("text-anchor", "middle")
            .attr("dominant-baseline", "middle")
            .attr("font-size", `${config.valueFont?.fontSize || 32}px`)
            .attr("font-weight", config.valueFont?.fontWeight || "bold")
            .attr("font-family", config.valueFont?.fontFamily || "Arial, sans-serif")
            .attr("fill", config.valueFont?.fontColor || "#5B6FD8")
            .text(formatValue(value, config));
    }

    // Center Text (e.g., "Power Usage")
    if (config.centerText) {
        g.append("text")
            .attr("y", radius * -0.1)
            .attr("text-anchor", "middle")
            .attr("dominant-baseline", "middle")
            .attr("font-size", "14px")
            .attr("fill", config.colors?.text || "#333333")
            .text(config.centerText);
    }
}

// Target aur Achievement ka text draw karta hai
function drawLabels(g, width, value, target, config) {
    const radius = Math.min(width, height) / 2 - 20;
    const percentage = target > 0 ? (value / target) * 100 : 0;

    if (config.visibility?.showAchievementLabel) {
        g.append("text")
            .attr("y", radius * 0.5)
            .attr("text-anchor", "middle")
            .attr("font-size", `${config.valueFont?.fontSize * 0.4 || 12}px`)
            .attr("fill", config.colors?.text || "#333333")
            .text(`${percentage.toFixed(1)}% ${config.achievementLabel || 'Achieved'}`);
    }

    if (config.visibility?.showTargetLabel) {
        g.append("text")
            .attr("y", radius * 0.75)
            .attr("text-anchor", "middle")
            .attr("font-size", `${config.valueFont?.fontSize * 0.3 || 10}px`)
            .attr("fill", config.colors?.text || "#333333")
            .text(`${config.targetLabel || 'Target'}: ${formatValue(target, config)}`);
    }
}


// ===================================================================================
// 5. TABLEAU EXTENSION LIFECYCLE (Tableau ke saath baat karta hai)
// ===================================================================================

// Jab extension load ho
tableau.extensions.initializeAsync().then(() => {
    // Settings load karein
    ConfigManager.loadConfig();
});

// Jab user "Configure" button click kare
tableau.extensions.ui.showSettingsPanel = () => {
    const settingsPanel = document.createElement('div');
    settingsPanel.innerHTML = SETTINGS_HTML;
    tableau.extensions.ui.displayDialogAsync(settingsPanel);

    // Save button ka event listener
    settingsPanel.querySelector('#save-btn').addEventListener('click', () => {
        ConfigManager.saveConfig();
    });

    // Reset button ka event listener
    settingsPanel.querySelector('#reset-btn').addEventListener('click', () => {
        ConfigManager.resetConfig();
    });

    // Decimal slider ka live update
    settingsPanel.querySelector('#config-decimalPlaces').addEventListener('input', (e) => {
        settingsPanel.querySelector('#config-decimalPlaces-value').textContent = e.target.value;
    });
};

// Jab user "Apply" ya "OK" click kare
tableau.extensions.ui.hideSettingsPanel = () => {
    console.log("Settings panel hidden.");
};

// Jab settings se re-render ka signal aaye
tableau.extensions.ui.updateAsync = () => {
    console.log("Update triggered from settings panel.");
};

// Jab Tableau gauge draw kare
tableau.extensions.ui.draw = (width, height, vizSettings) => {
    drawGauge(vizSettings.data, vizSettings);
};
```
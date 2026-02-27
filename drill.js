/* global d3 */
/* global tableau */

// --- 1. CONFIGURATION & STATE ---
const barWidth = 200; 
const barHeight = 20;
const margin = { top: 40, right: 120, bottom: 40, left: 100 };
const palette = ['#4e79a7', '#f28e2c', '#e15759', '#76b7b2', '#59a14f', '#edc949', '#af7aa1', '#ff9da7', '#9c755f', '#bab0ab'];
const colorScale = d3.scaleOrdinal().range(palette);

let svg, g, root, i = 0;
let width, height, dx, dy;
const duration = 500; 

// --- 2. AUTO-SETUP HTML CONTAINERS (Fix for "Container Not Found" error) ---
function setupDOM() {
    // Check if container exists, if not, create it
    if (!document.getElementById('chart-container')) {
        const div = document.createElement('div');
        div.id = 'chart-container';
        // Basic Styles to ensure it fills the screen
        div.style.width = '100%';
        div.style.height = '100%';
        div.style.position = 'absolute';
        div.style.top = '0';
        div.style.left = '0';
        document.body.appendChild(div);
    }

    // Check if tooltip exists, if not, create it
    if (!document.getElementById('tooltip')) {
        const tip = document.createElement('div');
        tip.id = 'tooltip';
        tip.style.position = 'absolute';
        tip.style.opacity = '0';
        tip.style.background = 'white';
        tip.style.border = '1px solid #ddd';
        tip.style.padding = '10px';
        tip.style.borderRadius = '4px';
        tip.style.pointerEvents = 'none';
        tip.style.zIndex = '10';
        document.body.appendChild(tip);
    }
}

// --- 3. TABLEAU INITIALIZATION ---
window.onload = function () {
    // Run DOM Setup first
    setupDOM();

    tableau.extensions.initializeAsync().then(() => {
        console.log("Tableau Extension Initialized");

        // 1. Initialize SVG
        initChart();

        // 2. Listen for data changes
        const worksheet = tableau.extensions.worksheetContent.worksheet;
        worksheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, loadData);

        // 3. Initial Load
        loadData();

        // 4. Resize Listener
        window.addEventListener('resize', () => {
            updateSize();
            if(root) update(root);
        });

    }).catch(err => {
        console.error("Error:", err);
        document.body.innerHTML = "<div style='color:red;padding:20px;'>Error: " + err.message + "</div>";
    });
};

// --- 4. CHART SETUP ---
function initChart() {
    const container = document.getElementById('chart-container');
    width = container.clientWidth;
    height = container.clientHeight;
    dx = 80;  // Vertical gap badhayein (Text aur Bar ke liye)
    dy = 300; // Horizontal gap badhayein (Bar width + gap)

    svg = d3.select("#chart-container").append("svg")
        .attr("width", width)
        .attr("height", height)
        .attr("viewBox", [0, 0, width, height])
        .style("font", "12px sans-serif")
        .style("user-select", "none");

    g = svg.append("g")
        .attr("transform", `translate(${margin.left},${height / 2})`);
}

function updateSize() {
    const container = document.getElementById('chart-container');
    width = container.clientWidth;
    height = container.clientHeight;
    dy = width / 6;
    svg.attr("width", width).attr("height", height).attr("viewBox", [0, 0, width, height]);
    g.attr("transform", `translate(${margin.left},${height / 2})`);
}

// --- 5. DATA FETCHING & PROCESSING ---
async function loadData() {
    try {
        const worksheet = tableau.extensions.worksheetContent.worksheet;
        const dataTableReader = await worksheet.getSummaryDataReaderAsync(undefined, { ignoreSelection: true });
        let rawData = [];

        for (let currentPage = 0; currentPage < dataTableReader.pageCount; currentPage++) {
            const page = await dataTableReader.getPageAsync(currentPage);
            rawData = rawData.concat(parseTableauPage(page));
        }
        await dataTableReader.releaseAsync();

        const { dims, measure } = analyzeData(rawData);

        if (dims.length === 0) {
            console.warn("No dimensions found.");
            return; 
        }

        const hierarchyData = buildHierarchy(rawData, dims, measure);
        renderChart(hierarchyData);

    } catch (error) {
        console.error("Data Fetch Error:", error);
    }
}

function parseTableauPage(page) {
    const rows = [];
    const columns = page.columns;
    const data = page.data;
    for (let i = 0; i < data.length; ++i) {
        const row = {};
        for (let j = 0; j < columns.length; ++j) {
            const colName = columns[j].fieldName;
            const val = data[i][columns[j].index];
            if (val && typeof val === 'object') {
                row[colName] = val.value;
                row[colName + '_formatted'] = val.formattedValue;
            } else {
                row[colName] = val;
            }
        }
        rows.push(row);
    }
    return rows;
}

function analyzeData(data) {
    if (data.length === 0) return { dims: [], measure: null };
    const keys = Object.keys(data[0]).filter(k => !k.includes('_formatted'));
    const dims = [];
    let measure = null;

    keys.forEach(key => {
        if (typeof data[0][key] === 'string') {
            dims.push(key);
        } else if (typeof data[0][key] === 'number' && !measure) {
            measure = key;
        }
    });
    return { dims, measure };
}

// --- 6. HIERARCHY BUILDER ---
function buildHierarchy(data, dims, measure) {
    const rootData = { name: "Total", children: [], value: 0 };
    const levelMap = new Map();
    levelMap.set("root_total", rootData);

    data.forEach(row => {
        let parentKey = "root_total";
        let parentPath = "Total";
        const rowValue = measure ? row[measure] : 1;

        for (let i = 0; i < dims.length; i++) {
            const dimName = dims[i];
            const dimValue = row[dimName] || "Unknown";
            const currentPath = parentPath + "_" + dimValue;
            
            if (!levelMap.has(currentPath)) {
                const newNode = {
                    name: dimValue,
                    type: dimName,
                    children: [],
                    value: 0
                };
                levelMap.set(currentPath, newNode);
                levelMap.get(parentKey).children.push(newNode);
            }

            levelMap.get(currentPath).value += rowValue;
            parentKey = currentPath;
            parentPath = currentPath;
        }
    });

    return rootData;
}

// --- 7. D3 RENDERING ---
function renderChart(data) {
    root = d3.hierarchy(data);
    
    // Collapse all initially
    root.descendants().forEach((d) => {
        if (d.depth > 0) { 
            d._children = d.children; 
            d.children = null; 
        }
    });

    root.x0 = height / 2;
    root.y0 = 0;
    update(root);
}

function update(source) {
    // Tree layout calculation
    const treemap = d3.tree().nodeSize([dx, dy]);
    treemap(root);
    
    const nodes = root.descendants();
    const links = root.links();

    // Horizontal position adjust karein
    nodes.forEach(d => { d.y = d.depth * dy; });

    // --- NODE UPDATE ---
    const node = g.selectAll("g.node")
        .data(nodes, d => d.id || (d.id = ++i));

    // 1. Enter New Nodes
    const nodeEnter = node.enter().append("g")
        .attr("class", "node")
        .attr("transform", d => `translate(${source.y0},${source.x0})`)
        .on("click", (event, d) => {
            // Toggle children on click
            if (d.children) {
                d._children = d.children;
                d.children = null;
            } else if (d._children) {
                d.children = d._children;
                d._children = null;
            }
            update(d);
        });

    // A. Gray Background Bar (Full Width)
    nodeEnter.append("rect")
        .attr("class", "bar-bg")
        .attr("rx", 5) // Rounded corners
        .attr("ry", 5)
        .attr("width", barWidth)
        .attr("height", 8) // Thin bar
        .attr("y", -4)
        .style("fill", "#f0f0f0");

    // B. Blue Value Bar (Percentage Width)
    nodeEnter.append("rect")
        .attr("class", "bar-fill")
        .attr("rx", 5)
        .attr("ry", 5)
        // Width calculation: Current Value / Root Total Value
        .attr("width", d => {
            const percent = (d.data.value / root.data.value); 
            return Math.max(0, percent * barWidth);
        })
        .attr("height", 8)
        .attr("y", -4)
        .style("fill", "#007bff"); // Blue color like image

    // C. Text: Name + Percentage (Top)
    nodeEnter.append("text")
        .attr("dy", "-10px")
        .attr("x", 0)
        .style("font-weight", "bold")
        .style("fill", "#333")
        .text(d => {
            const percent = ((d.data.value / root.data.value) * 100).toFixed(0);
            return `${d.data.name} (${percent}%)`;
        });

    // D. Text: Value (Bottom)
    nodeEnter.append("text")
        .attr("dy", "18px")
        .attr("x", 0)
        .style("font-size", "10px")
        .style("fill", "#777")
        .text(d => `Value: ${d3.format(",.0f")(d.data.value)}`);

    // E. Plus/Minus Icon (Right side)
    // Sirf unke liye jinke paas children hain
    const iconGroup = nodeEnter.append("g")
        .attr("class", "toggle-icon")
        .attr("transform", `translate(${barWidth + 15}, 0)`)
        .style("cursor", "pointer")
        .style("opacity", d => (d._children || d.children) ? 1 : 0);

    iconGroup.append("text")
        .attr("dy", "0.35em")
        .attr("text-anchor", "middle")
        .style("font-size", "18px")
        .style("font-weight", "bold")
        .text(d => d._children ? "+" : (d.children ? "-" : ""));

    // --- MERGE & UPDATE TRANSITIONS ---
    const nodeUpdate = nodeEnter.merge(node)
        .transition()
        .duration(duration)
        .attr("transform", d => `translate(${d.y},${d.x})`);

    // Update Icon State (+/-)
    nodeUpdate.select(".toggle-icon text")
        .text(d => d._children ? "+" : (d.children ? "-" : ""));
        
    // Update Icon Visibility
    nodeUpdate.select(".toggle-icon")
         .style("opacity", d => (d._children || d.children) ? 1 : 0);

    // --- EXIT TRANSITIONS ---
    const nodeExit = node.exit()
        .transition()
        .duration(duration)
        .attr("transform", d => `translate(${source.y},${source.x})`)
        .remove();
        
    nodeExit.select("rect").attr("width", 0);
    nodeExit.style("opacity", 0);

    // --- LINKS (Lines) ---
    const link = g.selectAll("path.link")
        .data(links, d => d.target.id);

    // Link Generation Logic
    // Source: Parent ke Bar ke RIGHT side se start hona chahiye
    // Target: Child ke Bar ke LEFT side pe end hona chahiye
    const linkGenerator = d3.linkHorizontal()
        .x(d => d.y)
        .y(d => d.x);

    // Custom Path Function to offset start point
    const diagonal = (d) => {
        const sourceX = d.source.y + barWidth + 25; // Bar width + gap space
        const sourceY = d.source.x;
        const targetX = d.target.y - 5; // Thoda pehle
        const targetY = d.target.x;
        
        return `M${sourceX},${sourceY}
                C${(sourceX + targetX) / 2},${sourceY}
                 ${(sourceX + targetX) / 2},${targetY}
                 ${targetX},${targetY}`;
    };

    const linkEnter = link.enter().insert("path", "g")
        .attr("class", "link")
        .attr("d", d => {
            const o = { x: source.x0, y: source.y0 };
            // Initial position trick for smooth animation
            return diagonal({ source: o, target: o }); 
        })
        .style("fill", "none")
        .style("stroke", "#ccc")
        .style("stroke-width", "1.5px");

    linkEnter.merge(link)
        .transition()
        .duration(duration)
        .attr("d", diagonal);

    link.exit()
        .transition()
        .duration(duration)
        .attr("d", d => {
            const o = { x: source.x, y: source.y };
            return diagonal({ source: o, target: o });
        })
        .remove();

    // Stash the old positions for transition.
    nodes.forEach(d => {
        d.x0 = d.x;
        d.y0 = d.y;
    });
}
// --- 8. TOOLTIPS ---
function handleMouseOver(event, d) {
    const tooltip = document.getElementById('tooltip');
    if (!tooltip) return;

    let content = `<strong>${d.data.name}</strong>`;
    if (d.data.type) content += `<br><span style="color:#666">Field: ${d.data.type}</span>`;
    if (d.data.value) content += `<br><span style="color:#666">Value: ${d3.format(',.0f')(d.data.value)}</span>`;

    tooltip.innerHTML = content;
    tooltip.style.opacity = 1;
    tooltip.style.left = (event.pageX + 15) + "px";
    tooltip.style.top = (event.pageY - 10) + "px";
    
    d3.select(this).select("circle").transition().duration(200).attr("r", 14);
}

function handleMouseOut(event, d) {
    const tooltip = document.getElementById('tooltip');
    if (!tooltip) return;
    
    tooltip.style.opacity = 0;
    d3.select(this).select("circle").transition().duration(200).attr("r", 10);
}
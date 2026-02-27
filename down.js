/* global d3 */
/* global tableau */

// --- 1. CONFIGURATION & STATE ---
const barWidth = 220;   
const barHeight = 25;   
const margin = { top: 40, right: 120, bottom: 40, left: 100 };

let svg, g, root, i = 0;
let width, height, dx, dy;
const duration = 500; 
let currentWorksheet; // <-- YEH LINE ADD KARO

// --- 2. AUTO-SETUP HTML CONTAINERS ---
function setupDOM() {
    if (!document.getElementById('chart-container')) {
        const div = document.createElement('div');
        div.id = 'chart-container';
        div.style.width = '100%';
        div.style.height = '100%';
        div.style.position = 'absolute';
        div.style.top = '0';
        div.style.left = '0';
        document.body.appendChild(div);
    }

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
        tip.style.boxShadow = '2px 2px 5px rgba(0,0,0,0.2)';
        document.body.appendChild(tip);
    }
}

// --- 3. TABLEAU INITIALIZATION ---
window.onload = function () {
    setupDOM();

    tableau.extensions.initializeAsync().then(() => {
        console.log("Tableau Extension Initialized");
        initChart();

        const worksheet = tableau.extensions.worksheetContent.worksheet;
        worksheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, loadData);

        loadData();

        window.addEventListener('resize', () => {
            updateSize();
            if(root) update(root);
        });

    }).catch(err => {
        console.error("Error:", err);
    });
};

// --- 4. CHART SETUP ---
function initChart() {
    const container = document.getElementById('chart-container');
    width = container.clientWidth;
    height = container.clientHeight;
    
    dx = 60;   
    dy = 350;  

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
    svg.attr("width", width).attr("height", height).attr("viewBox", [0, 0, width, height]);
    g.attr("transform", `translate(${margin.left},${height / 2})`);
}

// --- 5. DATA FETCHING & PROCESSING ---
async function loadData() {
    try {
        const worksheet = tableau.extensions.worksheetContent.worksheet;
        currentWorksheet = worksheet; // <-- YEH LINE ADD KARO
        const dataTableReader = await worksheet.getSummaryDataReaderAsync(undefined, { ignoreSelection: true });
        let rawData = [];

        for (let currentPage = 0; currentPage < dataTableReader.pageCount; currentPage++) {
            const page = await dataTableReader.getPageAsync(currentPage);
            rawData = rawData.concat(parseTableauPage(page));
        }
        await dataTableReader.releaseAsync();

        const { dims, measure } = analyzeData(rawData);
        if (dims.length === 0) return; 

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
    levelMap.set("Total", rootData);

    data.forEach(row => {
        let rawVal = measure ? row[measure] : 0;
        if (typeof rawVal === 'string') {
            rawVal = parseFloat(rawVal.replace(/,/g, ''));
        }
        const rowValue = isNaN(rawVal) ? 0 : rawVal;

        let currentPath = "Total";
        let currentNode = rootData;

        dims.forEach((dimName) => {
            let dimValue = row[dimName];
            if (dimValue === null || dimValue === undefined || dimValue === "") {
                dimValue = "Unknown";
            }

            const nextPath = currentPath + "/" + dimValue;

            if (!levelMap.has(nextPath)) {
                const newNode = {
                    name: String(dimValue),
                    type: dimName,
                    children: [],
                    value: 0
                };
                levelMap.set(nextPath, newNode);
                currentNode.children.push(newNode);
            }

            currentNode = levelMap.get(nextPath);
            currentPath = nextPath;
        });

        // Sirf leaf node mein value add karo
        currentNode.value += rowValue;
    });

    // ✅ CRITICAL FIX: Parent values bottom-up calculate karo
    function rollupValues(node) {
        if (node.children && node.children.length > 0) {
            node.children.forEach(child => rollupValues(child));
            // Parent = sum of all children
            node.value = node.children.reduce((sum, child) => sum + child.value, 0);
        }
        // Leaf node ki value already set hai - change mat karo
    }
    rollupValues(rootData);

    console.log('Hierarchy built:', JSON.stringify({
        name: rootData.name,
        value: rootData.value,
        children: rootData.children.map(c => ({name: c.name, value: c.value}))
    }, null, 2));

    return rootData;
}

// --- 7. D3 RENDERING ---
function renderChart(data) {
    g.selectAll("*").remove();

    root = d3.hierarchy(data);

    // ✅ KEY FIX: Manually assign d.value = d.data.value for all nodes
    // buildHierarchy has correct values in d.data.value
    // We copy them to d.value so bar calculations work correctly
    root.descendants().forEach(d => {
        d.value = d.data.value;
    });

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
    const treemap = d3.tree().nodeSize([dx, dy]);
    treemap(root);
    
    const nodes = root.descendants();
    const links = root.links();

    nodes.forEach(d => { d.y = d.depth * dy; });

    // --- NODES ---
    const node = g.selectAll("g.node")
        .data(nodes, d => d.id || (d.id = ++i));

    const nodeEnter = node.enter().append("g")
        .attr("class", "node")
        .attr("transform", d => `translate(${source.y0},${source.x0})`)
        .on("click", async (event, d) => {
    // Pehle expand/collapse karo (existing logic)
    if (d.children) {
        d._children = d.children;
        d.children = null;
    } else if (d._children) {
        d.children = d._children;
        d._children = null;
    }
    update(d);
    
    // Ab filter apply karo (naya logic)
    if (currentWorksheet && d.depth > 0) {
        await applyNodeFilter(d);
    } else if (currentWorksheet && d.depth === 0) {
        // Root (Total) pe click karne pe filter clear karo
        await currentWorksheet.clearFiltersAsync();
    }
})
        .on("mouseover", handleMouseOver)
        .on("mouseout", handleMouseOut);

    // 1. Gray Background Bar
    nodeEnter.append("rect")
        .attr("class", "bar-bg")
        .attr("rx", 4).attr("ry", 4)
        .attr("width", barWidth)
        .attr("height", barHeight)
        .attr("y", -barHeight / 2)
        .style("fill", "#f0f2f5")
        .style("stroke", "#e0e0e0");

    // 2. Blue Value Bar
    nodeEnter.append("rect")
        .attr("class", "bar-fill")
        .attr("rx", 4).attr("ry", 4)
        .attr("height", barHeight)
        .attr("y", -barHeight / 2)
        .attr("width", 0)
        .style("fill", "#007bff")
        .style("opacity", 0.85);

    // 3. ✅ FIX: Label = % relative to PARENT
    nodeEnter.append("text")
        .attr("class", "label-text")
        .attr("dy", "-18px")
        .attr("x", 2)
        .style("font-weight", "bold")
        .style("fill", "#333")
        .text(d => getLabel(d));

    // 4. Value Label
    nodeEnter.append("text")
        .attr("class", "value-text")
        .attr("dy", "28px")
        .attr("x", 2)
        .style("font-size", "12px")
        .style("fill", "#777")
        .text(d => `Value: ${d3.format(",.0f")(d.data.value)}`);

    // 5. Toggle Icon
    const iconGroup = nodeEnter.append("g")
        .attr("class", "toggle-icon")
        .attr("transform", `translate(${barWidth + 15}, 0)`)
        .style("cursor", "pointer");

    iconGroup.append("text")
        .attr("dy", "0.35em")
        .attr("text-anchor", "middle")
        .style("font-weight", "bold")
        .style("font-size", "16px")
        .text(d => d._children ? "+" : (d.children ? "-" : ""));

    // --- UPDATE TRANSITIONS ---
    const nodeUpdate = nodeEnter.merge(node)
        .transition()
        .duration(duration)
        .attr("transform", d => `translate(${d.y},${d.x})`);

    // ✅ FIX: Bar width = % of PARENT value (using d.value which is correctly set)
    nodeUpdate.select(".bar-fill")
        .attr("width", d => {
            const myVal    = d.value;                                   // this node
            const parentVal = d.parent ? d.parent.value : myVal;       // parent node
            if (!parentVal || parentVal === 0) return 0;
            const ratio = myVal / parentVal;
            return Math.max(2, ratio * barWidth);                       // min 2px so bar visible
        });

    nodeUpdate.select(".label-text")
        .text(d => getLabel(d));

    nodeUpdate.select(".value-text")
        .text(d => `Value: ${d3.format(",.0f")(d.data.value)}`);

    nodeUpdate.select(".toggle-icon")
        .style("opacity", d => (d.children || d._children) ? 1 : 0)
        .select("text")
        .text(d => d._children ? "+" : (d.children ? "-" : ""));

    // --- EXIT TRANSITIONS ---
    const nodeExit = node.exit()
        .transition()
        .duration(duration)
        .attr("transform", d => `translate(${source.y},${source.x})`)
        .remove();

    nodeExit.select("rect").attr("width", 0);
    nodeExit.style("opacity", 0);

    // --- LINKS ---
    const link = g.selectAll("path.link")
        .data(links, d => d.target.id);

    const diagonal = (d) => {
        const sX = d.source.y + barWidth + 30; 
        const sY = d.source.x;
        const tX = d.target.y - 10;
        const tY = d.target.x;
        return `M${sX},${sY}C${(sX + tX) / 2},${sY} ${(sX + tX) / 2},${tY} ${tX},${tY}`;
    };

    const linkEnter = link.enter().insert("path", "g")
        .attr("class", "link")
        .attr("d", d => {
             const fakeSource = {x: source.x0, y: source.y0};
             return diagonal({source: fakeSource, target: fakeSource});
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
            const fakeSource = {x: source.x, y: source.y};
            return diagonal({source: fakeSource, target: fakeSource});
        })
        .remove();

    nodes.forEach(d => {
        d.x0 = d.x;
        d.y0 = d.y;
    });
}

// ✅ KEY FUNCTION: % ALWAYS relative to direct parent
function getLabel(d) {
    if (!d.parent) {
        return `${d.data.name} (100%)`;
    }
    const parentVal = d.parent.data.value;
    if (!parentVal || parentVal === 0) return d.data.name;
    const percent = ((d.data.value / parentVal) * 100).toFixed(1);
    return `${d.data.name} (${percent}%)`;
}

// --- 8. TOOLTIPS ---
function handleMouseOver(event, d) {
    const tooltip = document.getElementById('tooltip');
    if (!tooltip) return;

    const parentVal = d.parent ? d.parent.data.value : d.data.value;
    const percentOfParent = parentVal > 0 ? ((d.data.value / parentVal) * 100).toFixed(1) : 100;
    const percentOfTotal  = root.data.value > 0 ? ((d.data.value / root.data.value) * 100).toFixed(1) : 100;

    let content = `<strong>${d.data.name}</strong>`;
    if (d.data.type) content += `<br><span style="color:#666">Field: ${d.data.type}</span>`;
    content += `<br><span style="color:#333">Value: ${d3.format(',.0f')(d.data.value)}</span>`;
    content += `<br><span style="color:#007bff">% of Parent (${d.parent ? d.parent.data.name : 'Total'}): ${percentOfParent}%</span>`;
    content += `<br><span style="color:#28a745">% of Total: ${percentOfTotal}%</span>`;

    tooltip.innerHTML = content;
    tooltip.style.opacity = 1;
    tooltip.style.left = (event.pageX + 15) + "px";
    tooltip.style.top = (event.pageY - 10) + "px";
    
    d3.select(event.currentTarget).select(".bar-fill").style("fill", "#0056b3"); 
}

function handleMouseOut(event, d) {
    const tooltip = document.getElementById('tooltip');
    if (!tooltip) return;
    tooltip.style.opacity = 0;
    d3.select(event.currentTarget).select(".bar-fill").style("fill", "#007bff"); 
}
// ✅ YEH FUNCTION YAHAN ADD KARO
async function applyNodeFilter(node) {
    try {
        // Collect all dimension filters from root to this node
        const filters = [];
        let current = node;
        
        // Traverse up to collect all dimension values
        while (current && current.depth > 0) {
            if (current.data.type && current.data.name) {
                filters.push({
                    fieldName: current.data.type,
                    values: [current.data.name]
                });
            }
            current = current.parent;
        }
        
                // ✅ SAARI DASHBOARD WORKSHEETS KO FILTER KARO
        const dashboard = tableau.extensions.dashboardContent.dashboard;
        const worksheets = dashboard.worksheets;
        
        // Pehle sabhi worksheets se filters clear karo
        for (const ws of worksheets) {
            await ws.clearFiltersAsync();
        }
        
        // Ab naya filter sabhi worksheets pe lagao
        for (const ws of worksheets) {
            for (const filter of filters) {
                try {
                    await ws.applyFilterAsync(
                        filter.fieldName,
                        filter.values,
                        tableau.FilterUpdateType.Replace
                    );
                } catch (e) {
                    // Agar kisi worksheet mein field nahi hai toh error ignore karo
                    console.log(`Filter not applied to ${ws.name}: ${e.message}`);
                }
            }
        }
        
        console.log("Dashboard filters applied to all worksheets:", filters);
        
    } catch (err) {
        console.error("Filter error:", err);
    }
}
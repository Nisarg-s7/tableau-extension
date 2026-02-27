tableau.extensions.initializeAsync().then(() => {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  const worksheet = dashboard.worksheets[0]; // first sheet

  loadData(worksheet);

  worksheet.addEventListener(
    tableau.TableauEventType.FilterChanged,
    () => loadData(worksheet)
  );
});

function loadData(worksheet) {
  worksheet.getSummaryDataAsync().then(dataTable => {
    const data = [];

    const dimIndex = 0;  // first column = Dimension
    const measureIndex = 1; // second column = Measure

    dataTable.data.forEach(row => {
      data.push({
        name: row[dimIndex].formattedValue,
        value: +row[measureIndex].value
      });
    });

    drawChart(data);
  });
}

function drawChart(data) {
  d3.select("#chart").html(""); // clear old

  const width = 600;
  const height = 400;
  const margin = { top: 20, right: 20, bottom: 50, left: 50 };

  const svg = d3.select("#chart")
    .append("svg")
    .attr("width", width)
    .attr("height", height);

  const x = d3.scaleBand()
    .domain(data.map(d => d.name))
    .range([margin.left, width - margin.right])
    .padding(0.2);

  const y = d3.scaleLinear()
    .domain([0, d3.max(data, d => d.value)])
    .nice()
    .range([height - margin.bottom, margin.top]);

  // Bars
  svg.selectAll("rect")
    .data(data)
    .enter()
    .append("rect")
    .attr("x", d => x(d.name))
    .attr("y", d => y(d.value))
    .attr("width", x.bandwidth())
    .attr("height", d => y(0) - y(d.value))
    .attr("fill", "#4e79a7");

  // X Axis
  svg.append("g")
    .attr("transform", `translate(0,${height - margin.bottom})`)
    .call(d3.axisBottom(x))
    .selectAll("text")
    .attr("transform", "rotate(-30)")
    .style("text-anchor", "end");

  // Y Axis
  svg.append("g")
    .attr("transform", `translate(${margin.left},0)`)
    .call(d3.axisLeft(y));
}

# Continuing Implementation

## Completed ✅
1. Updated encoding extraction (bars/lines/unfavorable/tooltip)
2. Created card generation logic with chartType and isUnfavorable flags
3. Updated getTrendClass to invert colors for unfavorable metrics
4. Renamed loadBarChartsAsync → loadChartsAsync
5. Updated renderKPIs to choose between bar/line based on chartType

## Remaining TODOs ⚠️

### 1. Add renderLineChart function
Insert this function after renderBarChart (around line 764 in main.js):

```javascript
// Render line chart for metric
function renderLineChart(elementId, currentData, referenceData, metricName, dateFieldName, isPercentage, isUnfavorable) {
  const container = document.getElementById(elementId);
  if (!container) return;

  container.innerHTML = '';

  const width = container.clientWidth;
  const height = container.clientHeight || 150;
  const margin = { top: 5, right: 0, bottom: 20, left: 0 };

  const svg = d3.select(container)
    .append('svg')
    .attr('width', '100%')
    .attr('height', '100%')
    .attr('viewBox', `0 0 ${width} ${height}`)
    .attr('preserveAspectRatio', 'xMidYMid meet');

  const x = d3.scaleTime()
    .domain([d3.min(currentData, d => d.date), d3.max(currentData, d => d.date)])
    .range([margin.left, width - margin.right]);

  const maxVal = Math.max(
    d3.max(currentData, d => d.value) || 0,
    d3.max(referenceData || [], d => d.value) || 0
  );

  const y = d3.scaleLinear()
    .domain([0, maxVal])
    .range([height - margin.bottom, margin.top]);

  const line = d3.line()
    .x(d => x(d.date))
    .y(d => y(d.value))
    .curve(d3.curveMonotoneX);

  // Reference line (gray)
  if (referenceData && referenceData.length > 0) {
    svg.append('path')
      .datum(referenceData)
      .attr('fill', 'none')
      .attr('stroke', '#e2e8f0')
      .attr('stroke-width', 2)
      .attr('d', line);
  }

  // Current line (dark blue)
  const currentPath = svg.append('path')
    .datum(currentData)
    .attr('fill', 'none')
    .attr('stroke', '#1e40af')
    .attr('stroke-width', 2.5)
    .attr('d', line);

  // Animate
  const totalLength = currentPath.node().getTotalLength();
  currentPath
    .attr('stroke-dasharray', totalLength + ' ' + totalLength)
    .attr('stroke-dashoffset', totalLength)
    .transition()
    .duration(800)
    .ease(d3.easeQuadOut)
    .attr('stroke-dashoffset', 0);

  // Dots
  svg.selectAll('.dot-current')
    .data(currentData)
    .enter()
    .append('circle')
    .attr('class', 'dot-current')
    .attr('cx', d => x(d.date))
    .attr('cy', d => y(d.value))
    .attr('r', 0)
    .attr('fill', '#1e40af')
    .transition()
    .duration(400)
    .delay((d, i) => 800 + i * 15)
    .attr('r', 3);

  // Axis labels
  if (currentData.length > 0) {
   const formatDate = d3.timeFormat('%b %d');
    svg.append('text')
      .attr('x', 0)
      .attr('y', height - 5)
      .attr('font-size', '10px')
      .attr('fill', '#9ca3af')
      .text(formatDate(currentData[0].date));

    svg.append('text')
      .attr('x', width)
      .attr('y', height - 5)
      .attr('text-anchor', 'end')
      .attr('font-size', '10px')
      .attr('fill', '#9ca3af')
      .text(formatDate(currentData[currentData.length - 1].date));
  }

  // Tooltip overlays
  svg.selectAll('.line-overlay')
    .data(currentData)
    .enter()
    .append('circle')
    .attr('cx', d => x(d.date))
    .attr('cy', d => y(d.value))
    .attr('r', 8)
    .attr('fill', 'transparent')
    .attr('cursor', 'pointer')
    .on('mouseenter', (e, d) => {
      const index = currentData.indexOf(d);
      const refVal = referenceData ? referenceData[index]?.value : 0;
      showTooltipForBar(e, d.date, d.value, refVal, metricName, isPercentage, isUnfavorable);
    })
    .on('mouseleave', hideTooltip)
    .on('mousemove', (e) => {
      lastEvent = e;
      updateTooltipPosition();
    });
}
```

### 2. Update renderBarChart signature
Change line 615 from:
```javascript
function renderBarChart(elementId, currentData, referenceData, metricName, dateFieldName, isPercentage) {
```

To:
```javascript
function renderBarChart(elementId, currentData, referenceData, metricName, dateFieldName, isPercentage, isUnfavorable) {
```

### 3. Update tooltips to use isUnfavorable
In `generateBarTooltipContent` and `generateTooltipContent`, invert colors based on isUnfavorable flag.

### 4. Add tooltip fields support
In tooltips, append extra fields from `metric.tooltipFields` at the end.

### 5. Update computeStateHash
Update to include new encodings in hash calculation.

## Next Steps
1. Manually add renderLineChart to main.js after line 764
2. Test in Tableau Desktop
3. Fix any remaining issues

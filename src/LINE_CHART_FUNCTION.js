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

    // X scale (time)
    const x = d3.scaleTime()
        .domain([d3.min(currentData, d => d.date), d3.max(currentData, d => d.date)])
        .range([margin.left, width - margin.right]);

    // Y scale
    // Y scale
    const allValues = [
        ...currentData.map(d => d.value),
        ...(referenceData || []).map(d => d.value)
    ];
    const minDataVal = d3.min(allValues) || 0;
    const maxDataVal = d3.max(allValues) || 0;

    // If all values are positive, we extend the upper bound to "squash" the chart downwards
    // so it doesn't look like it's floating too high.
    // If there are negative values, we let it scale naturally to fit them.
    let yDomain;
    if (minDataVal >= 0) {
        yDomain = [0, maxDataVal * 1.35]; // 1.35 factor pushes the chart down
    } else {
        // Add a little padding for negative values too
        const range = maxDataVal - minDataVal;
        yDomain = [minDataVal - range * 0.05, maxDataVal + range * 0.05];
    }

    const y = d3.scaleLinear()
        .domain(yDomain)
        .range([height - margin.bottom, margin.top]);

    // Line generators
    const line = d3.line()
        .x(d => x(d.date))
        .y(d => y(d.value))
        .curve(d3.curveMonotoneX);

    // Draw reference line (gray)
    if (referenceData && referenceData.length > 0) {
        svg.append('path')
            .datum(referenceData)
            .attr('fill', 'none')
            .attr('stroke', '#e2e8f0')
            .attr('stroke-width', 2)
            .attr('d', line);

        // Reference dots
        svg.selectAll('.dot-ref')
            .data(referenceData)
            .enter()
            .append('circle')
            .attr('class', 'dot-ref')
            .attr('cx', d => x(d.date))
            .attr('cy', d => y(d.value))
            .attr('r', 2)
            .attr('fill', '#e2e8f0');
    }

    // Draw current period line (dark blue)
    const currentPath = svg.append('path')
        .datum(currentData)
        .attr('fill', 'none')
        .attr('stroke', '#1e40af')
        .attr('stroke-width', 2.5)
        .attr('d', line);

    // Animate line drawing
    const totalLength = currentPath.node().getTotalLength();
    currentPath
        .attr('stroke-dasharray', totalLength + ' ' + totalLength)
        .attr('stroke-dashoffset', totalLength)
        .transition()
        .duration(800)
        .ease(d3.easeQuadOut)
        .attr('stroke-dashoffset', 0);

    // Current dots
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

    // Axis Labels (Start and End Date)
    if (currentData.length > 0) {
        const startDate = currentData[0].date;
        const endDate = currentData[currentData.length - 1].date;
        const formatDate = d3.timeFormat('%b %d');

        svg.append('text')
            .attr('x', 0)
            .attr('y', height - 5)
            .attr('font-size', '10px')
            .attr('fill', '#9ca3af')
            .text(formatDate(startDate));

        svg.append('text')
            .attr('x', width)
            .attr('y', height - 5)
            .attr('text-anchor', 'end')
            .attr('font-size', '10px')
            .attr('fill', '#9ca3af')
            .text(formatDate(endDate));
    }

    // Overlay for tooltips
    svg.selectAll('.line-overlay')
        .data(currentData)
        .enter()
        .append('circle')
        .attr('class', 'line-overlay')
        .attr('cx', d => x(d.date))
        .attr('cy', d => y(d.value))
        .attr('r', 8)
        .attr('fill', 'transparent')
        .attr('cursor', 'pointer')
        .on('mouseenter', (e, d) => {
            const index = currentData.indexOf(d);
            const refVal = referenceData ? referenceData[index]?.value : 0;
            showTooltipForBar(e, d.date, d.value, refVal, metricName, isPercentage, isUnfavorable);

            // Highlight dot
            const chartContainer = document.getElementById(elementId);
            const dots = chartContainer.querySelectorAll('.dot-current');
            if (dots[index]) {
                d3.select(dots[index])
                    .transition()
                    .duration(100)
                    .attr('r', 5)
                    .attr('stroke', '#1e40af')
                    .attr('stroke-width', 2);
            }
        })
        .on('mouseleave', () => {
            hideTooltip();
            const chartContainer = document.getElementById(elementId);
            const dots = chartContainer.querySelectorAll('.dot-current');
            dots.forEach(dot => {
                d3.select(dot)
                    .transition()
                    .duration(100)
                    .attr('r', 3)
                    .attr('stroke', 'none');
            });
        })
        .on('mousemove', (e) => {
            lastEvent = e;
            updateTooltipPosition();
        });
}

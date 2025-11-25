import './style.css';
import * as d3 from 'd3';

// -------------------- State --------------------
let state = {
  metrics: [],
  selectedPeriod: 'mtd',
  isCalculating: false,
  isApplyingOwnFilters: false,
  unregisterDataHandler: null,
  handleDataChange: null
};

// -------------------- Initialization --------------------
document.addEventListener('DOMContentLoaded', async () => {
  try {
    if (!window.tableau) {
      throw new Error('Tableau Extensions API is not loaded.');
    }
    await window.tableau.extensions.initializeAsync();
    console.log('Tableau Extension Initialized');
    const worksheet = window.tableau.extensions.worksheetContent.worksheet;

    // UI listeners
    document.getElementById('period-selector').addEventListener('change', e => {
      state.selectedPeriod = e.target.value;
      refreshKPIs(worksheet);
    });

    // Initialize tooltip container
    initTooltip();

    // Listen for data changes with debounce and filter check
    let resizeTimer;
    const handleDataChange = () => {
      // Ignore events triggered by our own filter changes
      if (state.isApplyingOwnFilters) {
        return;
      }

      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(() => {
        if (!state.isApplyingOwnFilters) {
          refreshKPIs(worksheet);
        }
      }, 1000); // Longer debounce to avoid rapid refreshes
    };

    state.unregisterDataHandler = worksheet.addEventListener(
      window.tableau.TableauEventType.SummaryDataChanged,
      handleDataChange
    );
    state.handleDataChange = handleDataChange;

    // Initial load
    await refreshKPIs(worksheet);

    // Right-click to reload data
    document.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      const menu = document.createElement('div');
      menu.style = `position:fixed; left:${e.pageX}px; top:${e.pageY}px; background:white; border:1px solid #ddd; padding:10px 14px; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,0.15); z-index:10001; cursor:pointer; font:13px Inter,sans-serif; color:#333`;
      menu.textContent = '🔄 Reload Extension';
      menu.onclick = async () => {
        menu.remove();
        const worksheet = window.tableau.extensions.worksheetContent.worksheet;
        await refreshKPIs(worksheet);
      };
      document.body.appendChild(menu);
      setTimeout(() => document.addEventListener('click', () => menu.remove(), { once: true }), 10);
    });
  } catch (err) {
    console.error('Initialization failed:', err);
    showDebug('Init Error: ' + err.message);
  }
});

// -------------------- Core Logic --------------------
async function refreshKPIs(worksheet) {
  if (state.isCalculating) return;
  state.isCalculating = true;
  state.isApplyingOwnFilters = true;

  document.getElementById('kpi-container').innerHTML = '<div class="loading">Calculating...</div>';
  console.time('Total Refresh Time');

  // Unregister listener temporarily
  if (state.unregisterDataHandler) {
    state.unregisterDataHandler();
    state.unregisterDataHandler = null;
  }

  try {
    // 1. Get Encodings
    console.time('Get Encodings');
    let metricFields = [];
    let dateFieldName = null;

    if (typeof worksheet.getVisualSpecificationAsync === 'function') {
      const spec = await worksheet.getVisualSpecificationAsync();
      const encodings = (spec.marksSpecifications && spec.marksSpecifications[0]?.encodings) || [];

      console.log('🔍 Raw Encodings:', JSON.stringify(encodings, null, 2));
      showDebug(`🔍 Encodings: ${encodings.length} found`);

      const getFieldNames = id =>
        encodings.filter(e => e.id === id)
          .map(e => e.field?.name || e.field || e.fieldName)
          .filter(Boolean);

      metricFields = getFieldNames('metric');
      const dateFields = getFieldNames('date');
      dateFieldName = dateFields[0] || null;

      showDebug(`🔍 Metrics (Spec): ${metricFields.join(', ')}`);
      showDebug(`🔍 Date (Spec): ${dateFieldName}`);
    }

    if (!dateFieldName) {
      const filters = await worksheet.getFiltersAsync();
      const dateFilter = filters.find(f => f.columnType === 'continuous-date' || f.columnType === 'discrete-date' || f.fieldName.toLowerCase().includes('date'));
      if (dateFilter) {
        dateFieldName = dateFilter.fieldName;
        showDebug(`🔍 Date (Filter): ${dateFieldName}`);
      }
    }
    console.timeEnd('Get Encodings');

    if (!dateFieldName) {
      showDebug('⚠️ No Date field found');
      return;
    }

    if (metricFields.length === 0) {
      // Fallback: Try to guess metrics from summary data columns if visual spec failed
      const summary = await worksheet.getSummaryDataAsync({ maxRows: 1 });
      console.log('🔍 Summary Columns:', summary.columns.map(c => c.fieldName));
      showDebug(`🔍 Summary Cols: ${summary.columns.length}`);

      const potentialMetrics = summary.columns
        .filter(c => c.dataType === 'float' || c.dataType === 'integer')
        .map(c => c.fieldName)
        .filter(name => !name.toLowerCase().includes('date') && !name.toLowerCase().includes('latitude') && !name.toLowerCase().includes('longitude'));

      if (potentialMetrics.length > 0) {
        metricFields = potentialMetrics;
        showDebug(`⚠️ Using fallback metrics: ${metricFields.join(', ')}`);
      } else {
        showDebug('⚠️ No Metrics found. Checked Spec and Summary.');
        showDebug(`Cols: ${summary.columns.map(c => c.fieldName).join(', ')}`);
        return;
      }
    }

    // 2. Define Periods
    const anchorDate = new Date();
    const periods = {
      current: getRange(state.selectedPeriod, anchorDate),
      prevMonth: getPrevMonthRange(getRange(state.selectedPeriod, anchorDate)),
      prevYear: getPrevYearRange(getRange(state.selectedPeriod, anchorDate))
    };

    // 3. Fetch Data (Sequential due to Tableau Filter API)
    console.time('Fetch Data');
    const results = {};

    const fetchDataForRange = async (rangeLabel, range) => {
      try {
        console.time(`Filter ${rangeLabel}`);
        await worksheet.applyRangeFilterAsync(dateFieldName, {
          min: range.start,
          max: range.end
        });
        console.timeEnd(`Filter ${rangeLabel}`);

        console.time(`GetData ${rangeLabel}`);
        const summary = await worksheet.getSummaryDataAsync({ ignoreSelection: true });
        console.timeEnd(`GetData ${rangeLabel}`);

        const data = summary.data;
        const columns = summary.columns;

        console.time(`Process ${rangeLabel}`);
        const values = {};
        metricFields.forEach(mName => { values[mName] = { val: 0, fmt: '' }; });

        // Optimize loop: Pre-calculate column indices
        const colIndices = {};
        metricFields.forEach(mName => {
          const col = columns.find(c => c.fieldName === mName);
          if (col) colIndices[mName] = col.index;
        });

        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          for (const mName of metricFields) {
            const idx = colIndices[mName];
            if (idx !== undefined) {
              const val = row[idx].nativeValue;
              if (typeof val === 'number') {
                values[mName].val += val;
              }
              if (!values[mName].fmt && row[idx].formattedValue) {
                values[mName].fmt = row[idx].formattedValue;
              }
            }
          }
        }
        console.timeEnd(`Process ${rangeLabel}`);

        results[rangeLabel] = values;
      } catch (e) {
        showDebug(`❌ Error: ${e.message}`);
        throw e;
      }
    };

    await fetchDataForRange('current', periods.current);
    await fetchDataForRange('prevMonth', periods.prevMonth);
    await fetchDataForRange('prevYear', periods.prevYear);
    console.timeEnd('Fetch Data');

    // 4. Clear Filter
    await worksheet.clearFilterAsync(dateFieldName);

    // 5. Fetch Bar Chart Data for each metric
    console.time('Fetch Bar Charts');
    const metricsWithCharts = await Promise.all(metricFields.map(async (mName) => {
      const curObj = results.current?.[mName];
      const prevMObj = results.prevMonth?.[mName];

      const curVal = curObj?.val || 0;
      const refVal = prevMObj?.val || 0;

      // Fetch bar chart data for current period (MTD)
      const barChartDataCurrent = await fetchBarChartData(
        worksheet,
        dateFieldName,
        mName,
        periods.current
      );

      // Fetch bar chart data for reference period (previous month)
      const barChartDataReference = await fetchBarChartData(
        worksheet,
        dateFieldName,
        mName,
        periods.prevMonth
      );

      return {
        name: mName,
        current: curVal,
        reference: refVal,
        prevMonth: prevMObj?.val || 0,
        prevYear: results.prevYear?.[mName]?.val || 0,
        isPercentage: curObj?.fmt?.includes('%') ?? false,
        formattedValue: curObj?.fmt,
        barChartDataCurrent,
        barChartDataReference,
        dateFieldName
      };
    }));
    console.timeEnd('Fetch Bar Charts');

    renderKPIs(metricsWithCharts);

  } catch (e) {
    console.error('Refresh Error:', e);
    showDebug('Refresh Error: ' + e.message);
  } finally {
    state.isCalculating = false;
    state.isApplyingOwnFilters = false;
    // Restore listener
    if (!state.unregisterDataHandler && state.handleDataChange) {
      state.unregisterDataHandler = worksheet.addEventListener(
        window.tableau.TableauEventType.SummaryDataChanged,
        state.handleDataChange
      );
    }
    console.timeEnd('Total Refresh Time');
  }
}

function renderKPIs(metrics) {
  const container = document.getElementById('kpi-container');
  container.innerHTML = '';

  metrics.forEach(metric => {
    const item = document.createElement('div');
    item.className = 'kpi-item';

    // Calculate deltas
    const momDiff = metric.current - metric.prevMonth;
    const momPct = metric.prevMonth ? (momDiff / metric.prevMonth) * 100 : 0;

    const chartId = `chart-${metric.name.replace(/\s+/g, '-')}`;

    item.innerHTML = `
      <div class="kpi-header">
        <div class="kpi-title">${metric.name}</div>
        <div class="kpi-value">${formatNumber(metric.current, metric.isPercentage)}</div>
      </div>
      <div class="kpi-comparison">
        <span class="${momDiff >= 0 ? 'positive' : 'negative'}">
          ${momDiff >= 0 ? '▲' : '▼'} ${Math.abs(momPct).toFixed(1)}%
        </span>
        <span class="text-muted">vs prev month</span>
      </div>
      <div id="${chartId}" class="bar-chart-container"></div>
    `;

    // Tooltip events
    item.addEventListener('mouseenter', (e) => showTooltipForMetric(e, metric));
    item.addEventListener('mouseleave', hideTooltip);
    item.addEventListener('mousemove', (e) => showTooltipForMetric(e, metric));

    container.appendChild(item);

    // Render bar chart
    if (metric.barChartDataCurrent && metric.barChartDataCurrent.length > 0) {
      renderBarChart(chartId, metric.barChartDataCurrent, metric.barChartDataReference, metric.name, metric.dateFieldName);
    }
  });
}

async function fetchBarChartData(worksheet, dateFieldName, metricField, range) {
  try {
    const dataPoints = [];
    const startDate = new Date(range.start);
    const endDate = new Date(range.end);

    // Iterate through each day in the range
    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
      const dayStart = new Date(d);
      dayStart.setUTCHours(0, 0, 0, 0);
      const dayEnd = new Date(d);
      dayEnd.setUTCHours(23, 59, 59, 999);

      await worksheet.applyRangeFilterAsync(dateFieldName, {
        min: dayStart,
        max: dayEnd
      });

      const summary = await worksheet.getSummaryDataAsync({ ignoreSelection: true });
      const dateIndex = summary.columns.findIndex(c => c.fieldName === dateFieldName);
      const metricIndex = summary.columns.findIndex(c => c.fieldName === metricField);

      let dailyValue = 0;
      if (metricIndex !== -1) {
        summary.data.forEach(row => {
          const val = row[metricIndex].nativeValue;
          if (typeof val === 'number') dailyValue += val;
        });
      }

      dataPoints.push({
        date: new Date(dayStart),
        value: dailyValue
      });
    }

    // Clear filter after loop
    await worksheet.clearFilterAsync(dateFieldName);
    return dataPoints;

  } catch (e) {
    console.warn(`Error fetching chart data for ${metricField}:`, e);
    return [];
  }
}

function renderBarChart(elementId, currentData, referenceData, metricName, dateFieldName) {
  const container = document.getElementById(elementId);
  if (!container) return;

  const width = container.clientWidth;
  const height = 60; // Fixed height for sparkline
  const margin = { top: 5, right: 0, bottom: 5, left: 0 };

  const svg = d3.select(container)
    .append('svg')
    .attr('width', width)
    .attr('height', height);

  // X scale
  const x = d3.scaleBand()
    .domain(currentData.map(d => d.date))
    .range([margin.left, width - margin.right])
    .padding(0.2);

  // Y scale (based on max of both datasets to keep scale consistent)
  const maxVal = Math.max(
    d3.max(currentData, d => d.value) || 0,
    d3.max(referenceData || [], d => d.value) || 0
  );

  const y = d3.scaleLinear()
    .domain([0, maxVal])
    .range([height - margin.bottom, margin.top]);

  // Draw Reference Bars (Gray)
  if (referenceData && referenceData.length > 0) {
    // We assume referenceData matches currentData by index (day 1 vs day 1)
    // Or we should map by day index. For MTD, it's day 1 to day N.
    // Let's assume they are aligned by day of month.

    svg.selectAll('.bar-ref')
      .data(currentData) // Use currentData to drive x-axis
      .enter()
      .append('rect')
      .attr('class', 'bar-ref')
      .attr('x', d => x(d.date))
      .attr('width', x.bandwidth() + 2) // Slightly wider for background effect
      .attr('y', (d, i) => {
        const refVal = referenceData[i]?.value || 0;
        return y(refVal);
      })
      .attr('height', (d, i) => {
        const refVal = referenceData[i]?.value || 0;
        return y(0) - y(refVal);
      })
      .attr('fill', '#cbd5e1') // Slate-300
      .attr('transform', `translate(-1, 0)`); // Center the wider bar
  }

  // Draw Current Bars (Conditional Color)
  svg.selectAll('.bar-current')
    .data(currentData)
    .enter()
    .append('rect')
    .attr('class', 'bar-current')
    .attr('x', d => x(d.date))
    .attr('width', x.bandwidth())
    .attr('y', d => y(d.value))
    .attr('height', d => y(0) - y(d.value))
    .attr('fill', (d, i) => {
      const refVal = referenceData?.[i]?.value || 0;
      return d.value > refVal ? '#4f46e5' : '#d97706'; // Indigo-600 vs Amber-600
    });
}

// -------------------- Helpers --------------------
function getRange(period, anchorDate) {
  const year = anchorDate.getUTCFullYear();
  const month = anchorDate.getUTCMonth();
  const day = anchorDate.getUTCDate();

  let start, end;
  end = new Date(Date.UTC(year, month, day, 23, 59, 59, 999));

  if (period === 'mtd') {
    start = new Date(Date.UTC(year, month, 1, 0, 0, 0, 0));
  } else if (period === 'qtd') {
    const qStart = Math.floor(month / 3) * 3;
    start = new Date(Date.UTC(year, qStart, 1, 0, 0, 0, 0));
  } else if (period === 'ytd') {
    start = new Date(Date.UTC(year, 0, 1, 0, 0, 0, 0));
  } else if (period === 'rolling_30') {
    const endDateUTC = new Date(Date.UTC(year, month, day, 0, 0, 0, 0));
    start = new Date(endDateUTC);
    start.setDate(start.getDate() - 29);
  }

  return { start, end };
}

function getPrevMonthRange(range) {
  const start = new Date(range.start);
  const end = new Date(range.end);
  start.setUTCMonth(start.getUTCMonth() - 1);
  end.setUTCMonth(end.getUTCMonth() - 1);
  return { start, end };
}

function getPrevYearRange(range) {
  const start = new Date(range.start);
  const end = new Date(range.end);
  start.setUTCFullYear(start.getUTCFullYear() - 1);
  end.setUTCFullYear(end.getUTCFullYear() - 1);
  return { start, end };
}

function formatNumber(val, isPercentage) {
  if (val === undefined || val === null) return '-';
  if (isPercentage) {
    return (val * 100).toFixed(1) + '%';
  }
  if (val >= 1000000) return (val / 1000000).toFixed(1) + 'M';
  if (val >= 1000) return (val / 1000).toFixed(1) + 'K';
  return val.toFixed(1);
}

function showDebug(message) {
  const debugDiv = document.getElementById('debug-info');
  if (debugDiv) {
    debugDiv.style.display = 'block';
    debugDiv.innerHTML += `<div>${message}</div>`;
  }
  console.log(message);
}

// -------------------- Tooltip --------------------
let tooltip = null;
let tooltipCache = new Map();
let rafId = null;
let lastEvent = null;

function initTooltip() {
  tooltip = document.createElement('div');
  tooltip.id = 'kpi-tooltip';
  tooltip.className = 'kpi-tooltip hidden';
  tooltip.style.willChange = 'transform';
  tooltip.style.top = '0';
  tooltip.style.left = '0';
  document.body.appendChild(tooltip);

  document.addEventListener('mousemove', (e) => {
    if (tooltip.classList.contains('hidden')) return;
    lastEvent = e;
    if (!rafId) {
      rafId = requestAnimationFrame(updateTooltipPosition);
    }
  });
}

function updateTooltipPosition() {
  if (!lastEvent || tooltip.classList.contains('hidden')) {
    rafId = null;
    return;
  }
  const margin = 15;
  let left = lastEvent.pageX + margin;
  let top = lastEvent.pageY + margin;
  const winW = window.innerWidth;
  const winH = window.innerHeight;
  const tipW = 280;
  const tipH = 150;

  if (left + tipW > winW) left = lastEvent.pageX - tipW - margin;
  if (top + tipH > winH) top = lastEvent.pageY - tipH - margin;

  left = Math.max(0, left);
  top = Math.max(0, top);

  tooltip.style.transform = `translate3d(${left}px, ${top}px, 0)`;
  rafId = null;
}

function showTooltipForMetric(e, metric) {
  tooltip.innerHTML = generateTooltipContent(metric);
  tooltip.classList.remove('hidden');
  lastEvent = e;
  updateTooltipPosition();
}

function hideTooltip() {
  tooltip.classList.add('hidden');
  if (rafId) {
    cancelAnimationFrame(rafId);
    rafId = null;
  }
}

function generateTooltipContent(metric) {
  const currentRange = getRange(state.selectedPeriod, new Date());
  const prevMonthRange = getPrevMonthRange(currentRange);
  const prevYearRange = getPrevYearRange(currentRange);
  const momDiff = metric.current - metric.prevMonth;
  const momPct = metric.prevMonth ? (momDiff / metric.prevMonth) * 100 : 0;
  const yoyDiff = metric.current - metric.prevYear;
  const yoyPct = metric.prevYear ? (yoyDiff / metric.prevYear) * 100 : 0;
  const formatDate = (date) => date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

  const formatDelta = (diff, pct, isPercentage) => {
    const triangle = diff >= 0 ? '▲' : '▼';
    const sign = diff >= 0 ? '+' : '';
    const deltaValue = isPercentage ? `${sign}${(diff * 100).toFixed(1)} pp` : `${sign}${formatNumber(Math.abs(diff), false)}`;
    const pctStr = `${sign}${pct.toFixed(1)}%`;
    return `${triangle} ${pctStr} <span class="tooltip-divider">|</span> ${deltaValue}`;
  };

  return `
    <div class="tooltip-header">${metric.name}</div>
    <div class="tooltip-section">
      <div class="tooltip-main-value">${formatNumber(metric.current, metric.isPercentage)}</div>
      <div class="tooltip-row"><span class="tooltip-label">Period:</span><span class="tooltip-value">${formatDate(currentRange.start)} - ${formatDate(currentRange.end)}</span></div>
    </div>
    <div class="tooltip-divider"></div>
    <div class="tooltip-section">
      <div class="tooltip-comparison-header">vs Previous Month</div>
      <div class="tooltip-row"><span class="tooltip-label">Period:</span><span class="tooltip-value">${formatDate(prevMonthRange.start)} - ${formatDate(prevMonthRange.end)}</span></div>
      <div class="tooltip-row"><span class="tooltip-label">Value:</span><span class="tooltip-value">${formatNumber(metric.prevMonth, metric.isPercentage)}</span></div>
      <div class="tooltip-row"><span class="tooltip-label">Δ:</span><span class="tooltip-value ${momDiff >= 0 ? 'positive' : 'negative'}">${formatDelta(momDiff, momPct, metric.isPercentage)}</span></div>
    </div>
    <div class="tooltip-divider"></div>
    <div class="tooltip-section">
      <div class="tooltip-comparison-header">vs Previous Year</div>
      <div class="tooltip-row"><span class="tooltip-label">Period:</span><span class="tooltip-value">${formatDate(prevYearRange.start)} - ${formatDate(prevYearRange.end)}</span></div>
      <div class="tooltip-row"><span class="tooltip-label">Value:</span><span class="tooltip-value">${formatNumber(metric.prevYear, metric.isPercentage)}</span></div>
      <div class="tooltip-row"><span class="tooltip-label">Δ:</span><span class="tooltip-value ${yoyDiff >= 0 ? 'positive' : 'negative'}">${formatDelta(yoyDiff, yoyPct, metric.isPercentage)}</span></div>
    </div>
  `;
}

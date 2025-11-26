import './style.css';
import * as d3 from 'd3';

// -------------------- State --------------------
let state = {
  metrics: [],
  selectedPeriod: 'mtd',
  isCalculating: false,
  isApplyingOwnFilters: false,
  unregisterDataHandler: null,
  handleDataChange: null,
  lastStateHash: null // Hash to detect real changes
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
      // Reset hash to ensure refresh happens
      state.lastStateHash = null;
      refreshKPIs(worksheet);
    });

    // Initialize tooltip container
    initTooltip();

    // Listen for data changes with debounce and filter check
    let resizeTimer;
    const handleDataChange = async () => {
      // Ignore events triggered by our own filter changes
      if (state.isApplyingOwnFilters) {
        return;
      }

      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(async () => {
        if (!state.isApplyingOwnFilters) {
          // Check if there are real changes before refreshing
          const hasChanges = await checkForChanges(worksheet);
          if (hasChanges) {
            console.log('🔄 Changes detected, refreshing KPIs...');
            await refreshKPIs(worksheet);
          } else {
            console.log('⏭️ No changes detected, skipping refresh');
          }
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
})

// -------------------- Core Logic --------------------

// Check if there are real changes in the data that require a refresh
async function checkForChanges(worksheet) {
  try {
    const currentStateHash = await computeStateHash(worksheet);

    console.log('🔍 Current State Hash:', currentStateHash);
    console.log('🔍 Previous State Hash:', state.lastStateHash);

    // First load - always refresh
    if (state.lastStateHash === null) {
      state.lastStateHash = currentStateHash;
      return true;
    }

    // Compare with previous state
    const hasChanges = currentStateHash !== state.lastStateHash;
    if (hasChanges) {
      console.log('✅ State changed!');
      state.lastStateHash = currentStateHash;
    } else {
      console.log('⏭️ State unchanged');
    }

    return hasChanges;
  } catch (e) {
    console.error('Error checking for changes:', e);
    // If we can't check, assume there are changes to be safe
    return true;
  }
}


// Compute a hash of the current state (metrics, encodings, filters)
async function computeStateHash(worksheet) {
  let hashParts = [];

  try {
    // 1. Get encodings (metrics and date fields)
    if (typeof worksheet.getVisualSpecificationAsync === 'function') {
      const spec = await worksheet.getVisualSpecificationAsync();
      const encodings = (spec.marksSpecifications && spec.marksSpecifications[0]?.encodings) || [];

      // Extract metric and date fields
      const metricFields = encodings
        .filter(e => e.id === 'metric')
        .map(e => e.field?.name || e.field || e.fieldName)
        .filter(Boolean)
        .sort();

      const dateFields = encodings
        .filter(e => e.id === 'date')
        .map(e => e.field?.name || e.field || e.fieldName)
        .filter(Boolean)
        .sort();

      hashParts.push(`metrics:${metricFields.join(',')}`);
      hashParts.push(`dates:${dateFields.join(',')}`);
    }

    // 2. Get active filters
    const filters = await worksheet.getFiltersAsync();
    const filterHash = filters
      .map(f => {
        let filterStr = `${f.fieldName}:${f.filterType}`;

        // Add filter values/ranges for more precise detection
        if (f.filterType === 'range' && f.minValue !== undefined && f.maxValue !== undefined) {
          filterStr += `:${f.minValue}-${f.maxValue}`;
        } else if (f.appliedValues) {
          filterStr += `:${f.appliedValues.map(v => v.value).sort().join(',')}`;
        }

        return filterStr;
      })
      .sort()
      .join('|');

    hashParts.push(`filters:${filterHash}`);

    // 3. Add selected period
    hashParts.push(`period:${state.selectedPeriod}`);

    return hashParts.join('::');
  } catch (e) {
    console.error('Error computing state hash:', e);
    return Date.now().toString(); // Fallback to always refresh on error
  }
}

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
      } else {
        showDebug(`⚠️ Filters checked: ${filters.map(f => `${f.fieldName} (${f.columnType})`).join(', ')}`);
      }
    }
    console.timeEnd('Get Encodings');

    if (!dateFieldName) {
      showDebug('⚠️ No Date field found');
      const emptyState = document.getElementById('empty-state');
      emptyState.style.display = 'flex';
      document.getElementById('main-content').style.display = 'none';

      // Update empty state message if we have metrics but no date
      if (metricFields.length > 0) {
        emptyState.querySelector('div[style="font-weight: 500;"]').textContent = 'Missing Date Field';
        emptyState.querySelector('div[style="font-size: 12px; margin-top: 4px;"]').textContent = 'Please drag a Date field to the "Dates" box in the Marks card.';
      }
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
        document.getElementById('empty-state').style.display = 'flex';
        document.getElementById('main-content').style.display = 'none';
        return;
      }
    }

    // 2. Define Periods
    document.getElementById('empty-state').style.display = 'none';
    document.getElementById('main-content').style.display = 'flex';

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

    // 5. Prepare metrics data (without charts for now - lazy loading)
    const metricsData = metricFields.map((mName) => {
      const curObj = results.current?.[mName];
      const prevMObj = results.prevMonth?.[mName];

      const curVal = curObj?.val || 0;
      const refVal = prevMObj?.val || 0;

      return {
        name: mName,
        current: curVal,
        reference: refVal,
        prevMonth: prevMObj?.val || 0,
        prevYear: results.prevYear?.[mName]?.val || 0,
        isPercentage: curObj?.fmt?.includes('%') ?? false,
        formattedValue: curObj?.fmt,
        dateFieldName
      };
    });

    // Render KPIs immediately with skeleton charts
    renderKPIs(metricsData, true); // true = show skeleton

    // Lazy load bar charts in background
    loadBarChartsAsync(worksheet, dateFieldName, metricsData, periods);

    // Update state hash after successful refresh
    state.lastStateHash = await computeStateHash(worksheet);

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

function renderKPIs(metrics, showSkeleton = false) {
  const container = document.getElementById('kpi-container');
  container.innerHTML = '';

  metrics.forEach(metric => {
    const item = document.createElement('div');
    item.className = 'kpi-item';

    // Calculate deltas
    const momDiff = metric.current - metric.prevMonth;
    const momPct = metric.prevMonth ? (momDiff / metric.prevMonth) * 100 : 0;

    const yoyDiff = metric.current - metric.prevYear;
    const yoyPct = metric.prevYear ? (yoyDiff / metric.prevYear) * 100 : 0;

    const chartId = `chart-${metric.name.replace(/\s+/g, '-')}`;

    // Helper for trend class
    const getTrendClass = (val) => val >= 0 ? 'trend-up' : 'trend-down';
    const formatDelta = (val, isPct) => {
      const sign = val >= 0 ? '+' : '';
      return isPct ? `${sign}${val.toFixed(1)}%` : `${sign}${formatNumber(Math.abs(val), false)}`;
    };

    item.innerHTML = `
      <div class="big-value">${formatNumber(metric.current, metric.isPercentage)}</div>
      
      <div class="comparison-line">
        <div class="comp-item" title="Year over Year">
          <span class="comp-label">YoY:</span>
          <span class="comp-val ${getTrendClass(yoyDiff)}">
            ${yoyDiff >= 0 ? '▲' : '▼'} ${Math.abs(yoyPct).toFixed(1)}%
          </span>
          <span class="comp-divider">|</span>
          <span class="comp-val ${getTrendClass(yoyDiff)}">
             ${formatDelta(yoyDiff, metric.isPercentage)}
          </span>
        </div>
        
        <div class="comp-item" title="Month over Month">
          <span class="comp-label">MoM:</span>
          <span class="comp-val ${getTrendClass(momDiff)}">
            ${momDiff >= 0 ? '▲' : '▼'} ${Math.abs(momPct).toFixed(1)}%
          </span>
          <span class="comp-divider">|</span>
           <span class="comp-val ${getTrendClass(momDiff)}">
             ${formatDelta(momDiff, metric.isPercentage)}
          </span>
        </div>
      </div>

      <div class="metric-subtitle">${metric.name} ${state.selectedPeriod.toUpperCase()}</div>
      
      <div id="${chartId}" class="bar-chart-container" style="width: 100%; flex: 1; min-height: 100px; margin-top: 12px; display: flex; align-items: flex-end;"></div>
    `;

    // Tooltip events - ONLY on the big value
    const bigValueEl = item.querySelector('.big-value');
    if (bigValueEl) {
      bigValueEl.style.cursor = 'help'; // Indicate hoverable
      bigValueEl.addEventListener('mouseenter', (e) => showTooltipForMetric(e, metric));
      bigValueEl.addEventListener('mouseleave', hideTooltip);
      bigValueEl.addEventListener('mousemove', (e) => {
        lastEvent = e;
        updateTooltipPosition();
      });
    }

    container.appendChild(item);

    // Show skeleton or real chart
    if (showSkeleton) {
      renderSkeletonChart(chartId);
    } else if (metric.barChartDataCurrent && metric.barChartDataCurrent.length > 0) {
      renderBarChart(chartId, metric.barChartDataCurrent, metric.barChartDataReference, metric.name, metric.dateFieldName, metric.isPercentage);
    }
  });
}

// Render skeleton loading animation
function renderSkeletonChart(elementId) {
  const container = document.getElementById(elementId);
  if (!container) return;

  // Generate 20-25 random bars
  const barCount = 20 + Math.floor(Math.random() * 6);
  const bars = [];

  for (let i = 0; i < barCount; i++) {
    // Random height between 20% and 100%
    const height = 20 + Math.random() * 80;
    // Add slight delay to each bar's animation for wave effect
    const delay = (i * 0.05).toFixed(2);
    bars.push(`<div class="skeleton-bar" style="height: ${height}%; animation-delay: ${delay}s"></div>`);
  }

  container.innerHTML = `<div class="skeleton-chart">${bars.join('')}</div>`;
}

// Lazy load bar charts in background
async function loadBarChartsAsync(worksheet, dateFieldName, metrics, periods) {
  console.time('⏱️ Load Bar Charts (Async)');

  for (const metric of metrics) {
    try {
      console.log(`📊 Loading chart for ${metric.name}...`);

      // Fetch bar chart data for current period
      const barChartDataCurrent = await fetchBarChartData(
        worksheet,
        dateFieldName,
        metric.name,
        periods.current
      );

      // Fetch bar chart data for reference period
      const barChartDataReference = await fetchBarChartData(
        worksheet,
        dateFieldName,
        metric.name,
        periods.prevMonth
      );

      // Replace skeleton with real chart
      const chartId = `chart-${metric.name.replace(/\s+/g, '-')}`;
      if (barChartDataCurrent && barChartDataCurrent.length > 0) {
        renderBarChart(chartId, barChartDataCurrent, barChartDataReference, metric.name, dateFieldName, metric.isPercentage);
        console.log(`✅ Chart loaded for ${metric.name}`);
      }
    } catch (e) {
      console.error(`❌ Failed to load chart for ${metric.name}:`, e);
    }
  }

  console.timeEnd('⏱️ Load Bar Charts (Async)');
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

function renderBarChart(elementId, currentData, referenceData, metricName, dateFieldName, isPercentage) {
  const container = document.getElementById(elementId);
  if (!container) return;

  container.innerHTML = ''; // Clear previous chart

  const width = container.clientWidth;
  const height = container.clientHeight || 150; // Use container height or fallback
  const margin = { top: 5, right: 0, bottom: 20, left: 0 };

  const svg = d3.select(container)
    .append('svg')
    .attr('width', width)
    .attr('height', height);

  // X scale
  const x = d3.scaleBand()
    .domain(currentData.map(d => d.date))
    .range([margin.left, width - margin.right])
    .padding(0.2);

  // Y scale
  const maxVal = Math.max(
    d3.max(currentData, d => d.value) || 0,
    d3.max(referenceData || [], d => d.value) || 0
  );

  const y = d3.scaleLinear()
    .domain([0, maxVal])
    .range([height - margin.bottom, margin.top]);

  // Draw Reference Bars (Gray) - Full Width
  if (referenceData && referenceData.length > 0) {
    svg.selectAll('.bar-ref')
      .data(currentData)
      .enter()
      .append('rect')
      .attr('class', 'bar-ref')
      .attr('x', d => x(d.date))
      .attr('width', x.bandwidth()) // Full bandwidth
      .attr('y', (d, i) => {
        const refVal = referenceData[i]?.value || 0;
        return y(refVal);
      })
      .attr('height', (d, i) => {
        const refVal = referenceData[i]?.value || 0;
        return y(0) - y(refVal);
      })
      .attr('fill', '#e2e8f0'); // Lighter gray (Slate-200)
  }

  // Draw Current Bars (Conditional Color) - Half Width & Centered
  svg.selectAll('.bar-current')
    .data(currentData)
    .enter()
    .append('rect')
    .attr('class', 'bar-current')
    .attr('x', d => x(d.date) + x.bandwidth() * 0.25) // Center: 25% offset
    .attr('width', x.bandwidth() * 0.5) // 50% width
    .attr('y', height) // Start from bottom
    .attr('height', 0) // Start with 0 height
    .attr('fill', (d, i) => {
      const refVal = referenceData?.[i]?.value || 0;
      return d.value > refVal ? '#4f46e5' : '#d97706';
    })
    .transition() // Add entrance animation
    .duration(800)
    .delay((d, i) => i * 50) // More staggered delay
    .ease(d3.easeElasticOut.amplitude(1).period(0.8)) // Elastic bounce effect
    .attr('y', d => y(d.value))
    .attr('height', d => y(0) - y(d.value));

  // Axis Labels (Start and End Date)
  if (currentData.length > 0) {
    const startDate = currentData[0].date;
    const endDate = currentData[currentData.length - 1].date;
    const formatDate = d3.timeFormat('%b %d');

    svg.append('text')
      .attr('x', 0)
      .attr('y', height - 5)
      .attr('font-size', '10px')
      .attr('fill', '#9ca3af') // Gray-400
      .text(formatDate(startDate));

    svg.append('text')
      .attr('x', width)
      .attr('y', height - 5)
      .attr('text-anchor', 'end')
      .attr('font-size', '10px')
      .attr('fill', '#9ca3af')
      .text(formatDate(endDate));
  }

  // Overlay for Tooltips
  svg.selectAll('.bar-overlay')
    .data(currentData)
    .enter()
    .append('rect')
    .attr('class', 'bar-overlay')
    .attr('x', d => x(d.date))
    .attr('width', x.bandwidth())
    .attr('y', 0)
    .attr('height', height - margin.bottom)
    .attr('fill', 'transparent')
    .on('mouseenter', (e, d) => {
      const index = currentData.indexOf(d);
      const refVal = referenceData ? referenceData[index]?.value : 0;
      showTooltipForBar(e, d.date, d.value, refVal, metricName, isPercentage);

      // Highlight effect
      const chartContainer = document.getElementById(elementId);
      const bars = chartContainer.querySelectorAll('.bar-current');
      bars.forEach((bar, i) => {
        if (i === index) {
          bar.classList.add('active');
          bar.classList.remove('inactive');
        } else {
          bar.classList.add('inactive');
          bar.classList.remove('active');
        }
      });
    })
    .on('mouseleave', () => {
      hideTooltip();
      const chartContainer = document.getElementById(elementId);
      const bars = chartContainer.querySelectorAll('.bar-current');
      bars.forEach(bar => {
        bar.classList.remove('active');
        bar.classList.remove('inactive');
      });
    })
    .on('mousemove', (e) => {
      lastEvent = e;
      updateTooltipPosition();
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
    debugDiv.style.zIndex = '99999';
    debugDiv.style.background = 'rgba(255,255,255,0.8)';
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

function showTooltipForBar(e, date, currentVal, refVal, metricName, isPercentage) {
  tooltip.innerHTML = generateBarTooltipContent(date, currentVal, refVal, metricName, isPercentage);
  tooltip.classList.remove('hidden');
  lastEvent = e;
  updateTooltipPosition();
}

function generateBarTooltipContent(date, currentVal, refVal, metricName, isPercentage) {
  const diff = currentVal - refVal;
  const pct = refVal ? (diff / refVal) * 100 : 0;
  const formatDate = (d) => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

  const triangle = diff >= 0 ? '▲' : '▼';
  const sign = diff >= 0 ? '+' : '';
  const colorClass = diff >= 0 ? 'positive' : 'negative';

  const deltaValue = isPercentage ? `${sign}${(diff * 100).toFixed(1)} pp` : `${sign}${formatNumber(Math.abs(diff), false)}`;
  const pctStr = `${sign}${pct.toFixed(1)}%`;

  return `
        <div class="tooltip-header">${metricName}</div>
        <div class="tooltip-section">
            <div class="tooltip-row">
                <span class="tooltip-label">Date:</span>
                <span class="tooltip-value">${formatDate(date)}</span>
            </div>
             <div class="tooltip-divider"></div>
            <div class="tooltip-row">
                <span class="tooltip-label">Current:</span>
                <span class="tooltip-value">${formatNumber(currentVal, isPercentage)}</span>
            </div>
            <div class="tooltip-row">
                <span class="tooltip-label">Reference:</span>
                <span class="tooltip-value">${formatNumber(refVal, isPercentage)}</span>
            </div>
             <div class="tooltip-divider"></div>
            <div class="tooltip-row">
                <span class="tooltip-label">Δ:</span>
                <span class="tooltip-value ${colorClass}">
                    ${triangle} ${pctStr} <span class="tooltip-divider">|</span> ${deltaValue}
                </span>
            </div>
        </div>
    `;
}

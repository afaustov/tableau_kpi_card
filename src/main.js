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
  lastStateHash: null, // Hash to detect real changes
  chartCache: {} // Cache for chart data to avoid re-fetching
};


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

// -------------------- Initialization --------------------
document.addEventListener('DOMContentLoaded', async () => {
  try {
    if (!window.tableau) {
      throw new Error('Tableau Extensions API is not loaded.');
    }
    await window.tableau.extensions.initializeAsync();
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
            await refreshKPIs(worksheet);
          } else {
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
    // showDebug('Init Error: ' + err.message);
  }
})

// -------------------- Core Logic --------------------

// Check if there are real changes in the data that require a refresh
async function checkForChanges(worksheet) {
  try {
    const currentStateHash = await computeStateHash(worksheet);


    // First load - always refresh
    if (state.lastStateHash === null) {
      state.lastStateHash = currentStateHash;
      return true;
    }

    // Compare with previous state
    const hasChanges = currentStateHash !== state.lastStateHash;
    if (hasChanges) {
      state.lastStateHash = currentStateHash;
    } else {
    }

    return hasChanges;
  } catch (e) {
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

      // Extract fields from new encodings
      const orderedEncodings = encodings
        .filter(e => e.id === 'bars' || e.id === 'lines')
        .map(e => `${e.field?.name || e.field || e.fieldName}(${e.id})`)
        .join(',');

      const unfavorableFields = encodings
        .filter(e => e.id === 'unfavorable')
        .map(e => e.field?.name || e.field || e.fieldName)
        .filter(Boolean)
        .sort();

      const tooltipFields = encodings
        .filter(e => e.id === 'tooltip')
        .map(e => e.field?.name || e.field || e.fieldName)
        .filter(Boolean)
        .sort();

      const dateFields = encodings
        .filter(e => e.id === 'date')
        .map(e => e.field?.name || e.field || e.fieldName)
        .filter(Boolean)
        .sort();

      hashParts.push(`metrics:${orderedEncodings}`);
      hashParts.push(`unfavorable:${unfavorableFields.join(',')}`);
      hashParts.push(`tooltip:${tooltipFields.join(',')}`);
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
    return Date.now().toString(); // Fallback to always refresh on error
  }
}

async function refreshKPIs(worksheet) {
  if (state.isCalculating) return;
  state.isCalculating = true;
  state.isApplyingOwnFilters = true;

  document.getElementById('kpi-container').innerHTML = '<div class="loading">Calculating...</div>';

  // Unregister listener temporarily
  if (state.unregisterDataHandler) {
    state.unregisterDataHandler();
    state.unregisterDataHandler = null;
  }

  try {
    // 1. Get Encodings
    let metricFields = [];
    let dateFieldName = null;

    if (typeof worksheet.getVisualSpecificationAsync === 'function') {
      const spec = await worksheet.getVisualSpecificationAsync();
      const encodings = (spec.marksSpecifications && spec.marksSpecifications[0]?.encodings) || [];

      // showDebug(`🔍 Encodings: ${encodings.length} found`);

      const getFieldNames = id =>
        encodings.filter(e => e.id === id)
          .map(e => e.field?.name || e.field || e.fieldName)
          .filter(Boolean);

      // Get fields from new encodings
      const barsFields = getFieldNames('bars');
      const linesFields = getFieldNames('lines');
      const unfavorableFields = getFieldNames('unfavorable');
      const tooltipFields = getFieldNames('tooltip');

      // Create ordered metrics list
      const orderedMetrics = [];
      encodings.forEach(e => {
        const fieldName = e.field?.name || e.field || e.fieldName;
        if (!fieldName) return;

        if (e.id === 'bars') {
          orderedMetrics.push({ name: fieldName, type: 'bar' });
        } else if (e.id === 'lines') {
          orderedMetrics.push({ name: fieldName, type: 'line' });
        }
      });

      // Combine bars and lines as metricFields for data fetching
      metricFields = [...new Set(orderedMetrics.map(m => m.name))];

      const dateFields = getFieldNames('date');
      dateFieldName = dateFields[0] || null;

      // showDebug(`🔍 Ordered Metrics: ${orderedMetrics.map(m => `${m.name}(${m.type})`).join(', ')}`);
      // showDebug(`🔍 Unfavorable: ${unfavorableFields.join(', ')}`);
      // showDebug(`🔍 Tooltip: ${tooltipFields.join(', ')}`);
      // showDebug(`🔍 Date: ${dateFieldName}`);

      // Store encoding info in state for later use
      state.encodings = { barsFields, linesFields, unfavorableFields, tooltipFields, orderedMetrics };
    }

    if (!dateFieldName) {
      const filters = await worksheet.getFiltersAsync();
      const dateFilter = filters.find(f => f.columnType === 'continuous-date' || f.columnType === 'discrete-date' || f.fieldName.toLowerCase().includes('date'));
      if (dateFilter) {
        dateFieldName = dateFilter.fieldName;
        // showDebug(`🔍 Date (Filter): ${dateFieldName}`);
      } else {
        // showDebug(`⚠️ Filters checked: ${filters.map(f => `${f.fieldName} (${f.columnType})`).join(', ')}`);
      }
    }

    if (!dateFieldName) {
      // showDebug('⚠️ No Date field found');
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
      // showDebug(`🔍 Summary Cols: ${summary.columns.length}`);

      const potentialMetrics = summary.columns
        .filter(c => c.dataType === 'float' || c.dataType === 'integer')
        .map(c => c.fieldName)
        .filter(name => !name.toLowerCase().includes('date') && !name.toLowerCase().includes('latitude') && !name.toLowerCase().includes('longitude'));

      if (potentialMetrics.length > 0) {
        metricFields = potentialMetrics;
        // showDebug(`⚠️ Using fallback metrics: ${metricFields.join(', ')}`);
      } else {
        // showDebug('⚠️ No Metrics found. Checked Spec and Summary.');
        // showDebug(`Cols: ${summary.columns.map(c => c.fieldName).join(', ')}`);
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
    const results = {};

    const fetchDataForRange = async (rangeLabel, range) => {
      try {
        await worksheet.applyRangeFilterAsync(dateFieldName, {
          min: range.start,
          max: range.end
        });

        const summary = await worksheet.getSummaryDataAsync({ ignoreSelection: true });

        const data = summary.data;
        const columns = summary.columns;

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

        results[rangeLabel] = values;
      } catch (e) {
        // showDebug(`❌ Error: ${e.message}`);
        throw e;
      }
    };

    await fetchDataForRange('current', periods.current);
    await fetchDataForRange('prevMonth', periods.prevMonth);
    await fetchDataForRange('prevYear', periods.prevYear);

    // 4. Clear Filter
    await worksheet.clearFilterAsync(dateFieldName);

    // 5. Prepare metrics data - create separate cards for bars and lines
    const cards = [];

    function createCardData(mName, chartType) {
      const curObj = results.current?.[mName];
      const prevMObj = results.prevMonth?.[mName];
      const curVal = curObj?.val || 0;
      const refVal = prevMObj?.val || 0;
      const isUnfavorable = state.encodings.unfavorableFields.includes(mName);

      return {
        name: mName,
        current: curVal,
        reference: refVal,
        prevMonth: prevMObj?.val || 0,
        prevYear: results.prevYear?.[mName]?.val || 0,
        isPercentage: curObj?.fmt?.includes('%') ?? false,
        formattedValue: curObj?.fmt,
        dateFieldName,
        chartType, // 'bar' or 'line'
        isUnfavorable,
        tooltipFields: state.encodings.tooltipFields
      };
    }

    // Create cards preserving order
    if (state.encodings.orderedMetrics && state.encodings.orderedMetrics.length > 0) {
      state.encodings.orderedMetrics.forEach(metric => {
        cards.push(createCardData(metric.name, metric.type));
      });
    } else {
      // Fallback
      state.encodings.barsFields.forEach(mName => cards.push(createCardData(mName, 'bar')));
      state.encodings.linesFields.forEach(mName => cards.push(createCardData(mName, 'line')));
    }

    // Render KPIs immediately with skeleton charts
    renderKPIs(cards, true); // true = show skeleton

    // Lazy load charts in background
    loadChartsAsync(worksheet, dateFieldName, cards, periods);

    // Update state hash after successful refresh
    state.lastStateHash = await computeStateHash(worksheet);

  } catch (e) {
    // showDebug('Refresh Error: ' + e.message);
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

    const chartId = `chart-${metric.name.replace(/\s+/g, '-')}-${metric.chartType}`;

    // Helper for trend class - INVERTED for unfavorable metrics
    const getTrendClass = (val) => {
      if (metric.isUnfavorable) {
        // Inverted: negative is good (blue), positive is bad (orange)
        return val >= 0 ? 'trend-down' : 'trend-up';
      }
      return val >= 0 ? 'trend-up' : 'trend-down';
    };

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

    // Show skeleton or real chart based on chartType
    if (showSkeleton) {
      if (metric.chartType === 'line') {
        renderSkeletonLineChart(chartId);
      } else {
        renderSkeletonChart(chartId);
      }
    } else if (metric.chartDataCurrent && metric.chartDataCurrent.length > 0) {
      if (metric.chartType === 'line') {
        renderLineChart(chartId, metric.chartDataCurrent, metric.chartDataReference, metric.name, metric.dateFieldName, metric.isPercentage, metric.isUnfavorable);
      } else {
        renderBarChart(chartId, metric.chartDataCurrent, metric.chartDataReference, metric.name, metric.dateFieldName, metric.isPercentage, metric.isUnfavorable);
      }
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


// Render skeleton loading animation for line chart
function renderSkeletonLineChart(elementId) {
  const container = document.getElementById(elementId);
  if (!container) return;

  // Two gray lines simulating current and reference with pulse animation
  const svg = `
    <svg width="100%" height="100%" viewBox="0 0 100 100" preserveAspectRatio="none" style="overflow: visible;">
      <style>
        @keyframes pulse-stroke {
          0% { stroke: #e2e8f0; }
          50% { stroke: #94a3b8; }
          100% { stroke: #e2e8f0; }
        }
        .skeleton-line-main {
          animation: pulse-stroke 1.5s infinite ease-in-out;
        }
      </style>
      
      <!-- Reference line simulation (solid, thicker, wavy, lower) -->
      <path d="M0,85 Q10,92 20,85 Q30,78 40,85 Q50,92 60,85 Q70,78 80,85 Q90,92 100,85" 
            fill="none" 
            stroke="#e2e8f0" 
            stroke-width="1" />

      <!-- Current line simulation (solid, thicker, wavy, pulsing, lower) -->
      <path d="M0,95 Q10,80 20,90 Q30,98 40,85 Q50,75 60,90 Q70,98 80,85 Q90,75 100,90" 
            fill="none" 
            stroke="#cbd5e1" 
            stroke-width="1"
            class="skeleton-line-main" />
    </svg>
  `;

  container.innerHTML = `<div class="skeleton-chart" style="display:block; width:100%; height:100%;">${svg}</div>`;
}

// Lazy load charts (bars and lines) in background
async function loadChartsAsync(worksheet, dateFieldName, cards, periods) {

  for (const card of cards) {
    try {

      let chartDataCurrent, chartDataReference;
      const cacheKey = `${card.name}-${state.selectedPeriod}`;
      const cached = state.chartCache[cacheKey];

      // Check cache: Must match period AND total value (to detect global filter changes)
      // We use a small epsilon for float comparison
      const isCacheValid = cached &&
        Math.abs(cached.totalCurrent - card.current) < 0.01 &&
        Math.abs(cached.totalReference - card.reference) < 0.01;

      if (isCacheValid) {
        chartDataCurrent = cached.dataCurrent;
        chartDataReference = cached.dataReference;

        // Render both at once from cache
        const chartId = `chart-${card.name.replace(/\s+/g, '-')}-${card.chartType}`;
        if (chartDataCurrent && chartDataCurrent.length > 0) {
          if (card.chartType === 'line') {
            renderLineChart(chartId, chartDataCurrent, chartDataReference, card.name, dateFieldName, card.isPercentage, card.isUnfavorable);
          } else {
            renderBarChart(chartId, chartDataCurrent, chartDataReference, card.name, dateFieldName, card.isPercentage, card.isUnfavorable);
          }
        }
      } else {
        const chartId = `chart-${card.name.replace(/\s+/g, '-')}-${card.chartType}`;

        // Progressive loading: Fetch and render REFERENCE period first (gray bars)
        chartDataReference = await fetchBarChartData(
          worksheet,
          dateFieldName,
          card.name,
          periods.prevMonth
        );

        // Render reference period first (pass empty array for current)
        if (card.chartType === 'line') {
          renderLineChart(chartId, [], chartDataReference, card.name, dateFieldName, card.isPercentage, card.isUnfavorable);
        } else {
          renderBarChart(chartId, [], chartDataReference, card.name, dateFieldName, card.isPercentage, card.isUnfavorable);
        }

        // Fetch chart data for CURRENT period
        chartDataCurrent = await fetchBarChartData(
          worksheet,
          dateFieldName,
          card.name,
          periods.current
        );

        // Re-render with both current and reference data
        if (card.chartType === 'line') {
          renderLineChart(chartId, chartDataCurrent, chartDataReference, card.name, dateFieldName, card.isPercentage, card.isUnfavorable);
        } else {
          renderBarChart(chartId, chartDataCurrent, chartDataReference, card.name, dateFieldName, card.isPercentage, card.isUnfavorable);
        }

        // Update cache
        state.chartCache[cacheKey] = {
          totalCurrent: card.current,
          totalReference: card.reference,
          dataCurrent: chartDataCurrent,
          dataReference: chartDataReference
        };
      }
    } catch (e) {
    }
  }

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
    return [];
  }
}

function renderBarChart(elementId, currentData, referenceData, metricName, dateFieldName, isPercentage, isUnfavorable) {
  const container = document.getElementById(elementId);
  if (!container) return;

  // Clear previous chart
  container.innerHTML = '';

  // Use ResizeObserver to handle responsiveness
  if (!container._resizeObserver) {
    container._resizeObserver = new ResizeObserver(entries => {
      // Placeholder for resize logic if needed
    });
  }

  const width = container.clientWidth;
  const height = container.clientHeight || 150;
  const margin = { top: 5, right: 0, bottom: 20, left: 0 };

  const svg = d3.select(container)
    .append('svg')
    .attr('width', '100%')
    .attr('height', '100%')
    .attr('viewBox', `0 0 ${width} ${height}`)
    .attr('preserveAspectRatio', 'xMidYMid meet');

  const hasCurrent = currentData && currentData.length > 0;
  const hasRef = referenceData && referenceData.length > 0;

  // Determine which dataset drives the X-axis
  // If we have current data, use it (and overlay reference on top)
  // If we only have reference data (loading state), use reference
  const primaryData = hasCurrent ? currentData : (hasRef ? referenceData : []);

  if (primaryData.length === 0) return;

  // X scale
  const x = d3.scaleBand()
    .domain(primaryData.map(d => d.date))
    .range([margin.left, width - margin.right])
    .padding(0.2);

  // Y scale
  const maxVal = Math.max(
    d3.max(currentData || [], d => d.value) || 0,
    d3.max(referenceData || [], d => d.value) || 0
  );

  const y = d3.scaleLinear()
    .domain([0, maxVal])
    .range([height - margin.bottom, margin.top]);

  // Draw Reference Bars (Gray) - Full Width
  if (hasRef) {
    // If we have current data, we map reference data to current dates by index
    // If we don't have current data, we just draw reference data as is
    const refSource = hasCurrent ? currentData : referenceData;

    svg.selectAll('.bar-ref')
      .data(refSource)
      .enter()
      .append('rect')
      .attr('class', 'bar-ref')
      .attr('x', d => x(d.date))
      .attr('width', x.bandwidth())
      .attr('y', (d, i) => {
        // If overlaying, get value from referenceData by index
        // If standalone, get value from d (which is referenceData item)
        const val = hasCurrent ? (referenceData[i]?.value || 0) : d.value;
        return y(val);
      })
      .attr('height', (d, i) => {
        const val = hasCurrent ? (referenceData[i]?.value || 0) : d.value;
        return y(0) - y(val);
      })
      .attr('fill', '#e2e8f0');
  }

  // Draw Current Bars (Conditional Color) - Half Width & Centered
  if (hasCurrent) {
    svg.selectAll('.bar-current')
      .data(currentData)
      .enter()
      .append('rect')
      .attr('class', 'bar-current')
      .attr('x', d => x(d.date) + x.bandwidth() * 0.25)
      .attr('width', x.bandwidth() * 0.5)
      .attr('y', y(0))
      .attr('height', 0)
      .attr('fill', (d, i) => {
        const refVal = referenceData?.[i]?.value || 0;
        const isGrowth = d.value > refVal;
        const isGood = isUnfavorable ? !isGrowth : isGrowth;
        return isGood ? '#4f46e5' : '#ef4444';
      })
      .transition()
      .duration(400)
      .ease(d3.easeQuadOut)
      .attr('y', d => y(d.value))
      .attr('height', d => y(0) - y(d.value));
  }

  // Axis Labels (Start and End Date)
  if (primaryData.length > 0) {
    const startDate = primaryData[0].date;
    const endDate = primaryData[primaryData.length - 1].date;
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

  // --- Interaction Layer (Brush + Hover) ---
  setupBrushInteraction(
    svg,
    width,
    height,
    margin,
    x,
    currentData,
    referenceData,
    metricName,
    isPercentage,
    isUnfavorable,
    'bar',
    elementId
  );
}

// Render line chart for metric
// Render line chart for metric
function renderLineChart(elementId, currentData, referenceData, metricName, dateFieldName, isPercentage, isUnfavorable) {
  const container = document.getElementById(elementId);
  if (!container) return;

  container.innerHTML = '';
  container.style.overflow = 'hidden';

  const width = container.clientWidth;
  const height = container.clientHeight || 150;
  const margin = { top: 15, right: 10, bottom: 20, left: 10 };

  const svg = d3.select(container)
    .append('svg')
    .attr('width', '100%')
    .attr('height', '100%')
    .attr('viewBox', `0 0 ${width} ${height}`)
    .attr('preserveAspectRatio', 'xMidYMid meet');

  const hasCurrent = currentData && currentData.length > 0;
  const hasRef = referenceData && referenceData.length > 0;

  // Determine which dataset drives the X-axis
  const primaryData = hasCurrent ? currentData : (hasRef ? referenceData : []);

  if (primaryData.length === 0) return;

  // X scale (time)
  const x = d3.scaleTime()
    .domain(d3.extent(primaryData, d => d.date))
    .range([margin.left, width - margin.right]);

  // Y scale - smart domain calculation to fit all values
  const allValues = [
    ...(currentData || []).map(d => d.value),
    ...(referenceData || []).map(d => d.value)
  ];

  let minData = 0;
  let maxData = 100;

  if (allValues.length > 0) {
    minData = d3.min(allValues);
    maxData = d3.max(allValues);
  }

  const range = maxData - minData;
  const padding = range > 0 ? range * 0.15 : (Math.abs(maxData) * 0.15 || 10);

  const yMin = minData - padding;
  const yMax = maxData + padding;

  const y = d3.scaleLinear()
    .domain([yMin, yMax])
    .range([height - margin.bottom, margin.top]);

  // Line generator
  const line = d3.line()
    .x(d => x(d.date))
    .y(d => y(d.value))
    .curve(d3.curveMonotoneX);

  // Draw zero line (dashed)
  if (yMin <= 0 && yMax >= 0) {
    svg.append('line')
      .attr('x1', margin.left)
      .attr('x2', width - margin.right)
      .attr('y1', y(0))
      .attr('y2', y(0))
      .attr('stroke', '#9ca3af')
      .attr('stroke-width', 1)
      .attr('stroke-dasharray', '4,4')
      .attr('opacity', 0.5);
  }

  // Draw reference line (Solid Gray)
  if (hasRef) {
    let referenceLineData;

    if (hasCurrent) {
      // Map reference values to current dates (overlay)
      referenceLineData = currentData.map((d, i) => ({
        date: d.date,
        value: referenceData[i]?.value || 0
      }));
    } else {
      // Use reference data directly
      referenceLineData = referenceData;
    }

    svg.append('path')
      .datum(referenceLineData)
      .attr('fill', 'none')
      .attr('stroke', '#cbd5e1')
      .attr('stroke-width', 1.5)
      .attr('d', line);
  }

  // Draw current period line
  if (hasCurrent) {
    const currentPath = svg.append('path')
      .datum(currentData)
      .attr('fill', 'none')
      .attr('stroke', isUnfavorable ? '#ef4444' : '#4f46e5')
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

    // Hover Dot (initially hidden)
    const hoverDot = svg.append('circle')
      .attr('r', 4)
      .attr('fill', isUnfavorable ? '#ef4444' : '#4f46e5')
      .attr('stroke', '#fff')
      .attr('stroke-width', 2)
      .style('opacity', 0)
      .style('pointer-events', 'none');
  }

  // Axis labels
  const domain = x.domain();
  if (domain && domain.length >= 2) {
    const formatDate = d3.timeFormat('%b %d');
    svg.append('text')
      .attr('x', margin.left)
      .attr('y', height - 5)
      .attr('font-size', '10px')
      .attr('fill', '#9ca3af')
      .text(formatDate(domain[0]));

    svg.append('text')
      .attr('x', width - margin.right)
      .attr('y', height - 5)
      .attr('text-anchor', 'end')
      .attr('font-size', '10px')
      .attr('fill', '#9ca3af')
      .text(formatDate(domain[1]));
  }

  // --- Interaction Layer (Brush + Hover) ---
  setupBrushInteraction(
    svg,
    width,
    height,
    margin,
    x,
    currentData,
    referenceData,
    metricName,
    isPercentage,
    isUnfavorable,
    'line',
    elementId
  );
}

// -------------------- Interaction Logic (Brush & Hover) --------------------

function setupBrushInteraction(svg, width, height, margin, x, data, refData, metricName, isPct, isUnfavorable, chartType, elementId) {
  const brush = d3.brushX()
    .extent([[margin.left, 0], [width - margin.right, height - margin.bottom]])
    .on('start brush end', brushed);

  const brushGroup = svg.append('g')
    .attr('class', 'brush')
    .call(brush);

  // Custom event listeners for hover when NOT brushing
  brushGroup.selectAll('.overlay')
    .on('mousemove', function (event) {
      // If we are currently brushing (selection exists), don't do hover logic
      if (d3.brushSelection(this.parentNode)) return;

      const [mx] = d3.pointer(event);
      handleHover(mx);
    })
    .on('mouseleave', function () {
      if (d3.brushSelection(this.parentNode)) return;
      clearHover();
    });

  function brushed(event) {
    const selection = event.selection;

    if (selection) {
      // 1. Identify selected data points
      const [x0, x1] = selection;

      let selectedData = [];
      let selectedIndices = [];

      if (chartType === 'bar') {
        // For band scale, we check if the band center is within selection
        const step = x.step();
        data.forEach((d, i) => {
          const barCenter = x(d.date) + x.bandwidth() / 2;
          if (barCenter >= x0 && barCenter <= x1) {
            selectedData.push(d);
            selectedIndices.push(i);
          }
        });
      } else {
        // For time scale, simply invert
        const date0 = x.invert(x0);
        const date1 = x.invert(x1);
        data.forEach((d, i) => {
          if (d.date >= date0 && d.date <= date1) {
            selectedData.push(d);
            selectedIndices.push(i);
          }
        });
      }

      // 2. Highlight selected, dim others
      highlightSelection(selectedIndices);

      // 3. Show Aggregated Tooltip
      if (selectedData.length > 0) {
        updateAggregatedTooltip(event.sourceEvent, selectedData, selectedIndices, refData, metricName, isPct, isUnfavorable);
      } else {
        hideTooltip();
      }

    } else {
      // Selection cleared
      clearHighlight();
      hideTooltip();
    }
  }

  function handleHover(mx) {
    // Find closest data point
    let index = -1;

    if (chartType === 'bar') {
      // Find band index
      const domain = x.domain();
      const range = x.range();
      const eachBand = x.step();
      const indexRaw = Math.floor((mx - range[0]) / eachBand);
      if (indexRaw >= 0 && indexRaw < domain.length) {
        index = indexRaw;
      }
    } else {
      // Time scale bisect
      const date = x.invert(mx);
      const bisect = d3.bisector(d => d.date).center;
      index = bisect(data, date);
    }

    if (index !== -1 && data[index]) {
      const d = data[index];
      const refVal = refData ? refData[index]?.value : 0;

      // Show single tooltip
      // We pass a synthetic event or just use the current mouse position
      // Since we are inside mousemove, 'event' is available in the scope if we passed it, 
      // but here we need to rely on global 'lastEvent' or pass it down.
      // Let's rely on the fact that updateTooltipPosition uses lastEvent.
      // We need to update lastEvent manually if we want it to track perfectly, 
      // but showTooltipForBar sets lastEvent.
      // We'll pass a mock event with pageX/Y from the mousemove if needed, 
      // but actually showTooltipForBar expects the event to set lastEvent.

      // Hack: we need the original event to get pageX/Y.
      // Let's assume the caller passed it or we use d3.pointer
      // For simplicity, we'll just use the global 'lastEvent' which is updated by the document listener
      // OR we can pass the event from the mousemove handler above.

      // Let's update the mousemove handler to pass event to handleHover
      // (See modification in the listener above)

      // Actually, let's just use the standard tooltip logic
      showTooltipForBar(lastEvent, d.date, d.value, refVal, metricName, isPct, isUnfavorable);

      // Highlight single item
      highlightSelection([index]);

      // Show hover dot for line chart
      if (chartType === 'line') {
        const dot = svg.select('circle[fill="' + (isUnfavorable ? '#ef4444' : '#4f46e5') + '"]');
        if (!dot.empty()) {
          dot.attr('cx', x(d.date))
            .attr('cy', d3.select(svg.selectAll('path').nodes()[1]).attr('d').split('L') ?  // Complex to reverse engineer Y from path
              // Easier: re-calculate Y
              // We need the Y scale. It's not passed.
              // Alternative: Pass Y scale or re-calculate.
              // Let's pass Y scale? No, too many args.
              // We can select the dot and set opacity 1, but we need coordinates.
              // For now, let's skip the dot movement in this generic function 
              // OR assume we can find the dot and we know the data.
              // We know 'd.value', but we don't have 'y' scale here.
              // Let's just highlight the point by index if possible.
              // The line chart render function created a specific hoverDot.
              // Let's make the hoverDot accessible or move it here.
              // Actually, let's just re-implement the dot logic inside renderLineChart's scope?
              // No, we want shared logic.
              // Let's pass 'y' scale to this function.
              // Adding 'y' to args.
              0 : 0);
        }
      }
    }
  }

  function clearHover() {
    clearHighlight();
    hideTooltip();
  }

  function highlightSelection(indices) {
    const container = document.getElementById(elementId);

    if (chartType === 'bar') {
      const bars = container.querySelectorAll('.bar-current');
      bars.forEach((bar, i) => {
        if (indices.includes(i)) {
          bar.style.opacity = '1';
          bar.classList.add('active');
        } else {
          bar.style.opacity = '0.3';
          bar.classList.remove('active');
        }
      });
    } else {
      // Line chart: Highlight dots?
      // Line chart doesn't have individual bars. 
      // We can show dots for selected points.
      const dots = container.querySelectorAll('.dot-current');
      // If we don't have dots rendered by default (we optimized them away), we might need to add them.
      // In the optimized line chart, we removed individual dots for performance.
      // We can add them dynamically or just rely on the range highlight.

      // For line chart, maybe we just dim the line? No, that looks weird.
      // Let's just show the tooltip.
      // Or we can add a "highlight region" rect behind.
    }
  }

  function clearHighlight() {
    const container = document.getElementById(elementId);
    if (chartType === 'bar') {
      const bars = container.querySelectorAll('.bar-current');
      bars.forEach(bar => {
        bar.style.opacity = '1';
        bar.classList.remove('active');
      });
    }
  }
}

function updateAggregatedTooltip(event, selectedData, selectedIndices, refData, metricName, isPct, isUnfavorable) {
  // Calculate Aggregates
  const sumCurrent = d3.sum(selectedData, d => d.value);
  const sumRef = d3.sum(selectedIndices, i => refData ? (refData[i]?.value || 0) : 0);

  const diff = sumCurrent - sumRef;
  const pct = sumRef ? (diff / sumRef) * 100 : 0;

  const startDate = selectedData[0].date;
  const endDate = selectedData[selectedData.length - 1].date;

  const formatDate = d => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  const dateRangeStr = `${formatDate(startDate)} - ${formatDate(endDate)}`;
  const countStr = `${selectedData.length} days`;

  // Generate Content
  const triangle = diff >= 0 ? '▲' : '▼';
  const sign = diff >= 0 ? '+' : '';
  const colorClass = isUnfavorable
    ? (diff >= 0 ? 'negative' : 'positive')
    : (diff >= 0 ? 'positive' : 'negative');

  const deltaValue = isPct ? `${sign}${(diff * 100).toFixed(1)} pp` : `${sign}${formatNumber(Math.abs(diff), false)}`;
  const pctStr = `${sign}${pct.toFixed(1)}%`;

  const content = `
    <div class="tooltip-header">${metricName} (Selected)</div>
    <div class="tooltip-section">
        <div class="tooltip-row">
            <span class="tooltip-label">Range:</span>
            <span class="tooltip-value">${dateRangeStr}</span>
        </div>
        <div class="tooltip-row">
            <span class="tooltip-label">Count:</span>
            <span class="tooltip-value">${countStr}</span>
        </div>
         <div class="tooltip-divider"></div>
        <div class="tooltip-row">
            <span class="tooltip-label">Sum:</span>
            <span class="tooltip-value">${formatNumber(sumCurrent, isPct)}</span>
        </div>
        <div class="tooltip-row">
            <span class="tooltip-label">Ref Sum:</span>
            <span class="tooltip-value">${formatNumber(sumRef, isPct)}</span>
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

  tooltip.innerHTML = content;
  tooltip.classList.remove('hidden');

  // Update position based on the brush event or center of selection
  // For simplicity, use the mouse position from the event
  if (event) {
    lastEvent = event;
    updateTooltipPosition();
  }
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

  // Helper for color class
  const getColorClass = (diff) => {
    if (metric.isUnfavorable) {
      return diff >= 0 ? 'negative' : 'positive';
    }
    return diff >= 0 ? 'positive' : 'negative';
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
      <div class="tooltip-row"><span class="tooltip-label">Δ:</span><span class="tooltip-value ${getColorClass(momDiff)}">${formatDelta(momDiff, momPct, metric.isPercentage)}</span></div>
    </div>
    <div class="tooltip-divider"></div>
    <div class="tooltip-section">
      <div class="tooltip-comparison-header">vs Previous Year</div>
      <div class="tooltip-row"><span class="tooltip-label">Period:</span><span class="tooltip-value">${formatDate(prevYearRange.start)} - ${formatDate(prevYearRange.end)}</span></div>
      <div class="tooltip-row"><span class="tooltip-label">Value:</span><span class="tooltip-value">${formatNumber(metric.prevYear, metric.isPercentage)}</span></div>
      <div class="tooltip-row"><span class="tooltip-label">Δ:</span><span class="tooltip-value ${getColorClass(yoyDiff)}">${formatDelta(yoyDiff, yoyPct, metric.isPercentage)}</span></div>
    </div>
  `;
}

function showTooltipForBar(e, date, currentVal, refVal, metricName, isPercentage, isUnfavorable = false) {
  tooltip.innerHTML = generateBarTooltipContent(date, currentVal, refVal, metricName, isPercentage, isUnfavorable);
  tooltip.classList.remove('hidden');
  lastEvent = e;
  updateTooltipPosition();
}

function generateBarTooltipContent(date, currentVal, refVal, metricName, isPercentage, isUnfavorable = false) {
  const diff = currentVal - refVal;
  const pct = refVal ? (diff / refVal) * 100 : 0;
  const formatDate = (d) => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

  const triangle = diff >= 0 ? '▲' : '▼';
  const sign = diff >= 0 ? '+' : '';
  // Invert colors for unfavorable metrics
  const colorClass = isUnfavorable
    ? (diff >= 0 ? 'negative' : 'positive')  // Inverted
    : (diff >= 0 ? 'positive' : 'negative'); // Normal

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

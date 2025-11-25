import './style.css';

// -------------------- State --------------------
let state = {
  metrics: [],
  selectedPeriod: 'mtd',
  isCalculating: false,
  isApplyingOwnFilters: false
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
    let rowFields = [];
    let columnFields = [];

    if (typeof worksheet.getVisualSpecificationAsync === 'function') {
      const spec = await worksheet.getVisualSpecificationAsync();
      const encodings = (spec.marksSpecifications && spec.marksSpecifications[0]?.encodings) || [];

      const getFieldNames = id =>
        encodings.filter(e => e.id === id)
          .map(e => e.field?.name || e.field || e.fieldName)
          .filter(Boolean);

      metricFields = getFieldNames('metric');
      const dateFields = getFieldNames('date');
      dateFieldName = dateFields[0] || null;

      // Read from rows and columns encodings
      rowFields = getFieldNames('rows');
      columnFields = getFieldNames('columns');

      console.log('📊 Encodings detected:', {
        metrics: metricFields,
        date: dateFieldName,
        rows: rowFields,
        columns: columnFields
      });
    }

    if (!dateFieldName) {
      const filters = await worksheet.getFiltersAsync();
      const dateFilter = filters.find(f => f.columnType === 'continuous-date' || f.columnType === 'discrete-date' || f.fieldName.toLowerCase().includes('date'));
      if (dateFilter) {
        dateFieldName = dateFilter.fieldName;
      }
    }
    console.timeEnd('Get Encodings');

    if (!dateFieldName) {
      showDebug('⚠️ No Date field found');
      state.isCalculating = false;
      state.isApplyingOwnFilters = false;
      return;
    }

    if (metricFields.length === 0) {
      // Fallback: Try to guess metrics from summary data columns if visual spec failed
      const summary = await worksheet.getSummaryDataAsync({ maxRows: 1 });
      const potentialMetrics = summary.columns
        .filter(c => c.dataType === 'float' || c.dataType === 'integer')
        .map(c => c.fieldName)
        .filter(name => !name.toLowerCase().includes('date') && !name.toLowerCase().includes('latitude') && !name.toLowerCase().includes('longitude'));

      if (potentialMetrics.length > 0) {
        metricFields = potentialMetrics;
        showDebug(`⚠️ Using fallback metrics: ${metricFields.join(', ')}`);
      } else {
        showDebug('⚠️ No Metrics found');
        state.isCalculating = false;
        state.isApplyingOwnFilters = false;
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

        // Use standard for loop for better performance on large arrays
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
      const prevYObj = results.prevYear?.[mName];

      const curVal = curObj?.val || 0;
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

// -------------------- Debugging --------------------
function showDebug(message) {
        const debugDiv = document.getElementById('debug-info');
        if (debugDiv) {
          debugDiv.style.display = 'block';
          if (!debugDiv.innerHTML.includes('v3.3')) {
            debugDiv.innerHTML = '<div style="font-weight:bold; color:#00ff00">🟢 Build v3.3 (UI Polish)</div>';
          }
          debugDiv.innerHTML += `<div>${message}</div>`;
        }
        console.log(message);
      }

// -------------------- Tooltip Optimization --------------------
let tooltip = null;
    let tooltipCache = new Map();
    let rafId = null;
    let lastEvent = null;

    function initTooltip() {
      // Create tooltip element once
      tooltip = document.createElement('div');
      tooltip.id = 'kpi-tooltip';
      tooltip.className = 'kpi-tooltip hidden';
      // Use transform for positioning to avoid layout thrashing
      tooltip.style.willChange = 'transform';
      tooltip.style.top = '0';
      tooltip.style.left = '0';
      document.body.appendChild(tooltip);

      // Global mousemove for positioning (only active when needed)
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

      // Simple boundary check (using window size)
      const winW = window.innerWidth;
      const winH = window.innerHeight;

      // Approximate tooltip size to avoid thrashing
      const tipW = 280;
      const tipH = 150;

      if (left + tipW > winW) left = lastEvent.pageX - tipW - margin;
      if (top + tipH > winH) top = lastEvent.pageY - tipH - margin;

      left = Math.max(0, left);
      top = Math.max(0, top);

      // Use transform instead of top/left for GPU acceleration
      tooltip.style.transform = `translate3d(${left}px, ${top}px, 0)`;

      rafId = null;
    }

    function showTooltipForMetric(e, metric) {
      if (!tooltipCache.has(metric.name)) {
        tooltipCache.set(metric.name, generateTooltipContent(metric));
      }
      tooltip.innerHTML = tooltipCache.get(metric.name);
      tooltip.classList.remove('hidden');

      // Initial position
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

      // Format delta with triangles and separator
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

    // -------------------- Performance Monitoring --------------------
    (function () {
      const fpsDiv = document.createElement('div');
      fpsDiv.style = 'position:fixed; top:0; left:0; background:rgba(0,0,0,0.7); color:#0f0; padding:4px 8px; font-family:monospace; font-size:12px; z-index:99999; pointer-events:none;';
      document.body.appendChild(fpsDiv);

      let frameCount = 0;
      let lastTime = performance.now();

      function updateFPS() {
        frameCount++;
        const now = performance.now();
        if (now - lastTime >= 1000) {
          const fps = Math.round((frameCount * 1000) / (now - lastTime));
          fpsDiv.textContent = `FPS: ${fps} | v3.3`;
          frameCount = 0;
          lastTime = now;
        }
        requestAnimationFrame(updateFPS);
      }
      requestAnimationFrame(updateFPS);
    })();

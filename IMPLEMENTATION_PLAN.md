# Implementation Plan: Enhanced KPI Cards

## Current State
- Uses single "metric" encoding
- Shows bar charts for all metrics
- Colors: blue (positive), orange (negative)

## Required Changes

### 1. Encoding Structure
```javascript
// OLD: getFieldNames('metric')
// NEW:
const barsFields = getFieldNames('bars');
const linesFields = getFieldNames('lines');
const unfavorableFields = getFieldNames('unfavorable');
const tooltipFields = getFieldNames('tooltip');
```

### 2. Metrics Processing
```javascript
// For each field, determine:
// - visualizationType: 'bar' | 'line'
// - isUnfavorable: boolean
// - tooltipFields: string[]

// If metric is in both bars AND lines:
// Create 2 separate card objects with same data but different chartType
```

### 3. Color Inversion for Unfavorable
```javascript
// Helper function:
function getTrendClass(val, isUnfavorable) {
  if (isUnfavorable) {
    return val >= 0 ? 'trend-down' : 'trend-up'; // INVERTED
  }
  return val >= 0 ? 'trend-up' : 'trend-down';
}
```

### 4. Line Chart Rendering
```javascript
// New function: renderLineChart(elementId, currentData, referenceData, ...)
// - Gray line for reference period
// - Dark blue line for current period
// - Same data aggregation as bars (daily)
```

### 5. Tooltip Enhancement
```javascript
// Add tooltip fields at the end of both tooltips:
// - Main metric tooltip (on big value hover)
// - Bar/Line chart tooltip (on chart hover)
```

## Implementation Order
1. Update computeStateHash to include new encodings
2. Update refreshKPIs to fetch bars/lines/unfavorable/tooltip
3. Create metric objects with chartType and isUnfavorable flags
4. Update renderKPIs to use getTrendClass with isUnfavorable
5. Implement renderLineChart function
6. Update tooltips to include extra fields

# KPI Card Viz Extension

This is a Tableau Viz Extension that displays a premium KPI card with:
- Total Value (MTD, Rolling)
- Comparisons (YoY, MoM, WoW)
- Trend Chart (Line/Bar)

## Setup

1. **Install Dependencies**:
   ```bash
   npm install
   ```

2. **Run the Extension**:
   ```bash
   npm run dev
   ```
   The extension will be available at `http://localhost:5173`.

3. **Use in Tableau**:
   - Open Tableau Desktop (2024.1 or later).
   - Connect to data.
   - In a Worksheet, look for the "Marks" card.
   - Change the Mark type to "Viz Extension" (or "Add Extension" if available in the dropdown).
   - Select the `manifest.trex` file located in this directory.
   - Drag a **Measure** to the "Metric" drop zone.
   - Drag a **Date** dimension to the "Date" drop zone.

## Features

- **Metric**: Displays the sum of the selected measure.
- **Date**: Filters and calculates trends based on the selected date field.
- **Period Selector**: Choose between MTD (Month to Date) or Rolling 7/30/90 days.
- **Comparisons**: Automatically calculates Year-over-Year, Month-over-Month, and Week-over-Week changes.
- **Chart**: Toggle between Line and Bar charts to see the trend.

## Development

- `src/main.js`: Main logic for data processing and rendering.
- `src/style.css`: Styling (Glassmorphism, Premium UI).
- `manifest.trex`: Tableau Extension Manifest.

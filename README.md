# Monday Morning Report - Stoa Group

A comprehensive, brand-aligned Monday Morning Report for Stoa Group with PDF export functionality.

## Features

- **Clean, Modern Design**: Uses Stoa Group brand colors and typography
- **Comprehensive Data Sections**:
  - Occupancy metrics
  - Leasing performance
  - Renewals & Collections
  - Rent analysis
  - Google Reviews
- **PDF Export**: One-click PDF export with proper formatting
- **Responsive Layout**: Works on all screen sizes
- **Brand Colors**:
  - Primary Green: #7e8a6b
  - Primary Grey: #757270
  - Secondary Green: #a6ad8a
  - Secondary Blue: #bdc2ce
  - Secondary Grey: #efeff1

## Usage

1. The report automatically loads data from your Domo datasets (`mmrData` and `googleReviews`)
2. Click "Export as PDF" to generate a PDF version of the report
3. The PDF will be automatically downloaded with the filename: `Monday_Morning_Report_YYYY-MM-DD.pdf`

## Data Sources

The report uses two datasets defined in `manifest.json`:

1. **mmrData**: Primary Monday Morning Report data including occupancy, leasing, renewals, and rents
2. **googleReviews**: Google review data for properties

## Sections

### Occupancy
- Total units by property
- Current occupancy percentage
- Move-ins and move-outs
- Net change calculations

### Leasing
- Current leased percentage
- Visit tracking
- Gross/canceled/denied leases
- Net leases
- Closing ratios
- Gain percentages

### Renewals & Collections
- T-12 expired leases
- T-12 renewed leases
- Renewal rates
- In-service units
- Delinquent units tracking

### Rents
- Occupied rent vs. budgeted rent
- Move-in rent analysis
- Percentage differences
- Visual indicators for positive/negative performance

### Google Reviews
- Average ratings by property
- Overall portfolio rating

## Customization

The report can be customized by modifying:
- `app.css`: Styling and brand colors
- `app.js`: Data processing logic and calculations
- `index.html`: Report structure and layout

## Technical Details

- Built with vanilla JavaScript for optimal performance
- Uses html2pdf.js for PDF generation
- Domo API integration for live data
- Responsive design with mobile support
- Print-optimized styling


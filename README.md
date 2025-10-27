# Excel Performance Bottleneck Test Tool

An Excel Add-in built to reproduce and measure a performance bottleneck in Excel's Office JavaScript API when updating rows in an Excel Table.

## Overview

This tool demonstrates a performance degradation in Excel for Windows version **16.0.19127.20314** (October) compared to version **16.0.19029.20244** (September). This degradation occurs when using the Office-JS API to update table rows. Frequent updates to large numeric tables are a common use case in financial organisations. The October version is approximately 12.5 times slower than the September version. 

By default this tool does the following 
1. Creates a table of 15,000 rows and 17 columns
2. Updates a random 10% of the rows with random data 
3. Repeats step 2 1,000 times. 

Timestamped log entries are written to the JavaScript console. Our test results are stored in [RESULTS.md](./RESULTS.md). A marked performance degradation is observed using Excel v16.0.19127.20314

## Features

- **Table Creation**: Generates large test tables with configurable row counts (default: 15,000 rows)
- **Performance Testing**: Measures execution time for updating a percentage of table rows over multiple iterations

## Technical Stack

- **TypeScript** - Type-safe development
- **React 18** - UI framework
- **Vite** - Fast build tool and dev server
- **Fluent UI v9** - Microsoft's design system
- **Office JavaScript API** - Excel integration

## Prerequisites

- Node.js (LTS version recommended)
- Microsoft Excel for Desktop
- npm

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd excel-performance-bottleneck
```

2. Install dependencies:
```bash
npm install
```

## Performance Test Execution

1. Install the desired version using [ODT](https://www.microsoft.com/en-us/download/details.aspx?id=49117)

2. Start the development server:
```bash
npm run dev
```
3. Start Excel
Open a new terminal window (as the development server is still running)
```bash
npm run start
```

### Stage 1: Prepare Test Table
1. Open the add-in task pane in Excel
2. Configure the number of rows (default: 15,000)
3. Click "Prepare Test Table"
4. The add-in creates a table named "TestTable" with 17 columns (1 key column + 16 data columns)

### Stage 2: Execute Performance Test
1. Configure test parameters:
   - **Percentage of rows to update**: 1-100%
   - **Number of iterations**: How many times to repeat the update
2. Click "Execute Performance Test"
3. The tool will update the specified percentage of rows for each iteration and measure execution time
4. View the results displayed in the task pane


## Performance Issue Details

This tool was created to demonstrate a performance bottleneck where updating table rows in Excel version 16.0.19127.20314 takes significantly longer than in version 16.0.19029.20244 when using the Office JavaScript API to update rows of an `Excel.Table`

## Results

See [RESULTS.md](./RESULTS.md) for detailed test results comparing Excel versions.

## Author

DiffusionData

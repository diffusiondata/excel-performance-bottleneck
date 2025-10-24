# Excel Performance Bottleneck Test Tool

An Excel Add-in built to reproduce and measure a performance regression in Excel's Office JavaScript API when updating rows in an Excel Table.

## Overview

This tool demonstrates a significant performance degradation observed in Excel version **16.0.19127.20314** (Windows) compared to the prior version **16.0.19029.20244** when using the Office-JS API to update table rows.

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
- Microsoft Excel (Desktop or Web)
- npm or yarn package manager

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

1. Install the desired version. Use ODT to install either verion 

1. Start the development server:
```bash
npm run dev
```
2. Start Excel
Open a new shell (as the development server is still running)
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

This tool was created to demonstrate a performance regression where updating table rows in Excel version 16.0.19127.20314 takes significantly longer than in version 16.0.19029.20244 when using the Office JavaScript API to update rows of an `Excel.Table`

## Results



## Author

DiffusionData

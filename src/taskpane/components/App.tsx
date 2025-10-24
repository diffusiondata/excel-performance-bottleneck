import * as React from "react";
import {
  Button,
  Input,
  Divider,
  MessageBar,
  MessageBarBody,
  MessageBarIntent,
  Spinner,
  Text,
} from "@fluentui/react-components";
import Progress from "./Progress";

/* global console, Excel, performance  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  rows: number;
  percentage: number;
  iterations: number;
  isCreatingTable: boolean;
  isRunningTest: boolean;
  message: string;
  messageType: MessageBarIntent;
  executionTime: number | null;
}

const TABLE_NAME = "TestTable";
const NUM_COLUMNS = 17;

export default class App extends React.Component<AppProps, AppState> {
  constructor(props: AppProps, context: AppState) {
    super(props, context);
    this.state = {
      rows: 15000,
      percentage: 10,
      iterations: 1000,
      isCreatingTable: false,
      isRunningTest: false,
      message: "",
      messageType: "info",
      executionTime: null,
    };
  }

  generateRandomData = (rows: number, cols: number): number[][] => {
    const data: number[][] = [];
    for (let i = 0; i < rows; i++) {
      const row: number[] = [];
      for (let j = 0; j < cols; j++) {
        row.push(Math.random() * 1000);
      }
      data.push(row);
    }
    return data;
  };

  shuffleArray = <T,>(array: T[]): T[] => {
    const shuffled = [...array];
    for (let i = shuffled.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    return shuffled;
  };

  prepareTest = async () => {
    const { rows } = this.state;

    if (rows <= 0) {
      this.setState({
        message: "Rows must be a positive integer",
        messageType: "error",
      });
      return;
    }

    this.setState({ isCreatingTable: true, message: "", executionTime: null });

    try {
      await Excel.run(async (context) => {
        // Delete existing table if it exists
        const tables = context.workbook.tables;
        tables.load("items/name");
        await context.sync();

        const existingTable = tables.items.find((t) => t.name === TABLE_NAME);
        if (existingTable) {
          existingTable.delete();
          await context.sync();
        }

        // Get the active worksheet
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Generate data: first column is keys (0, 1, 2, ...), rest are random
        const data: (number | string)[][] = [];

        // Header row
        const headers: string[] = ["Key"];
        for (let i = 1; i < NUM_COLUMNS; i++) {
          headers.push(`Col${i}`);
        }
        data.push(headers);

        // Data rows
        for (let i = 0; i < rows; i++) {
          const row: number[] = [i]; // Key column
          for (let j = 1; j < NUM_COLUMNS; j++) {
            row.push(Math.random() * 1000);
          }
          data.push(row);
        }

        // Create range and table
        const range = sheet.getRangeByIndexes(0, 0, rows + 1, NUM_COLUMNS);
        range.values = data;

        // Create table
        const table = sheet.tables.add(range, true);
        table.name = TABLE_NAME;

        await context.sync();

        this.setState({
          isCreatingTable: false,
          message: `Table '${TABLE_NAME}' created successfully with ${rows} rows and ${NUM_COLUMNS} columns`,
          messageType: "success",
        });
      });
    } catch (error) {
      console.error(error);
      this.setState({
        isCreatingTable: false,
        message: `Error creating table: ${error instanceof Error ? error.message : String(error)}`,
        messageType: "error",
      });
    }
  };

  executeTest = async () => {
    const { percentage, iterations } = this.state;

    if (percentage < 1 || percentage > 100) {
      this.setState({
        message: "Percentage must be between 1 and 100",
        messageType: "error",
      });
      return;
    }

    if (iterations <= 0) {
      this.setState({
        message: "Iterations must be a positive integer",
        messageType: "error",
      });
      return;
    }

    this.setState({ isRunningTest: true, message: "", executionTime: null });

    const startTime = performance.now();

    try {
      await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(TABLE_NAME);
        const dataRange = table.getDataBodyRange();
        dataRange.load("values, rowCount, columnCount");
        await context.sync();

        const totalRows = dataRange.rowCount;
        const numRowsToUpdate = Math.ceil((totalRows * percentage) / 100);

        // Generate all keys
        const allKeys = Array.from({ length: totalRows }, (_, i) => i);

        for (let iter = 0; iter < iterations; iter++) {
          // 1. Generate new random data for non-key columns
          const randomData = this.generateRandomData(numRowsToUpdate, NUM_COLUMNS - 1);

          // 2. Choose random set of rows
          const shuffledKeys = this.shuffleArray(allKeys);
          const selectedKeys = shuffledKeys.slice(0, numRowsToUpdate);

          // 3. Update the non-key values in the table
          for (let i = 0; i < selectedKeys.length; i++) {
            const rowIndex = selectedKeys[i];
            const rowRange = dataRange.getRow(rowIndex);

            // Update only non-key columns (columns 1 to 16)
            const updateRange = rowRange.getOffsetRange(0, 1).getAbsoluteResizedRange(1, NUM_COLUMNS - 1);
            // updateRange.load("address");
            // await context.sync();

            // console.debug(`Applying data to range ${updateRange.address}`, { data: randomData[i] });
            updateRange.values = [randomData[i]];
          }

          // Sync every iteration to ensure updates are applied (intentional for performance testing)
          // eslint-disable-next-line office-addins/no-context-sync-in-loop
          await context.sync();
        }
      });

      const endTime = performance.now();
      const executionTime = endTime - startTime;

      const message = `Test completed. ${iterations} iterations on ${percentage}% of rows`;
      console.log(message);
      this.setState({
        isRunningTest: false,
        message,
        messageType: "success",
        executionTime,
      });
    } catch (error) {
      console.error(error);
      const message = `Error executing test: ${error instanceof Error ? error.message : String(error)}`;
      console.error(message);
      this.setState({
        isRunningTest: false,
        message,
        messageType: "error",
      });
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;
    const { rows, percentage, iterations, isCreatingTable, isRunningTest, message, messageType, executionTime } =
      this.state;

    if (!isOfficeInitialized) {
      return <Progress title={title} logo="/icon-128.png" message="Please sideload your addin to see app body." />;
    }

    return (
      <div style={{ padding: "20px", display: "flex", flexDirection: "column", gap: "20px" }}>
        <h2>{title}</h2>

        {message && (
          <MessageBar intent={messageType}>
            <MessageBarBody>{message}</MessageBarBody>
          </MessageBar>
        )}

        {/* Form 1: Table Preparation */}
        <div style={{ display: "flex", flexDirection: "column", gap: "15px" }}>
          <h3>1. Prepare Test Table</h3>
          <div>
            <label htmlFor="rows-input" style={{ display: "block", marginBottom: "4px" }}>
              Rows
            </label>
            <Input
              id="rows-input"
              type="number"
              value={rows.toString()}
              onChange={(_, data) => this.setState({ rows: parseInt(data.value || "0", 10) })}
              disabled={isCreatingTable || isRunningTest}
            />
          </div>
          <Button appearance="primary" onClick={this.prepareTest} disabled={isCreatingTable || isRunningTest}>
            Prepare Test
          </Button>
          {isCreatingTable && <Spinner label="Creating table..." />}
        </div>

        <Divider />

        {/* Form 2: Test Execution */}
        <div style={{ display: "flex", flexDirection: "column", gap: "15px" }}>
          <h3>2. Execute Performance Test</h3>
          <div>
            <label htmlFor="percentage-input" style={{ display: "block", marginBottom: "4px" }}>
              Percentage (1-100)
            </label>
            <Input
              id="percentage-input"
              type="number"
              value={percentage.toString()}
              onChange={(_, data) => this.setState({ percentage: parseInt(data.value || "0", 10) })}
              disabled={isCreatingTable || isRunningTest}
            />
          </div>
          <div>
            <label htmlFor="iterations-input" style={{ display: "block", marginBottom: "4px" }}>
              Iterations
            </label>
            <Input
              id="iterations-input"
              type="number"
              value={iterations.toString()}
              onChange={(_, data) => this.setState({ iterations: parseInt(data.value || "0", 10) })}
              disabled={isCreatingTable || isRunningTest}
            />
          </div>
          <Button appearance="primary" onClick={this.executeTest} disabled={isCreatingTable || isRunningTest}>
            Execute Test
          </Button>
          {isRunningTest && <Spinner label="Running test..." />}
          {executionTime !== null && (
            <MessageBar intent="info">
              <MessageBarBody>
                <strong>Execution Time:</strong> {executionTime.toFixed(2)} ms ({(executionTime / 1000).toFixed(2)}{" "}
                seconds)
              </MessageBarBody>
            </MessageBar>
          )}
        </div>

        <Divider />

        <h2>Purpose</h2>
        <Text>This tool measures the performance of the the office-js API, when updating rows in an Excel.Table.</Text>
        <Text>
          A considerable drop in performance has been observed in version Excel 16.0.19127.20314 for Windows. Whereas
          the prior version (16.0.19029.20244) does not suffers the same performance problem.
        </Text>

        <h2>Use</h2>
        <Text>
          Prepare a table of 15,000 rows. Then choose a percentage of those rows to update, and the number of
          iterations. Click <code>Execute Test</code> and observe the table update. When the test is complete the
          duration is displayed and logged to the consol.e
        </Text>
      </div>
    );
  }
}

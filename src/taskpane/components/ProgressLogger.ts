import {logWithTime} from "./Common";
import prettyMilliseconds from "pretty-ms";

/**
 * For periodic progress logging, for long running jobs.
 */
export class ProgressLogger {
  private readonly total: number;
  private readonly intervalMs: number;
  private readonly startTime: number;
  private lastLogTime: number;
  private lastLogIteration: number;
  private currentIter: number;
  private timer: ReturnType<typeof setTimeout> | null;

  constructor(totalIterations: number, intervalMs: number = 60000) {
    this.total = totalIterations;
    this.intervalMs = intervalMs;
    this.startTime = performance.now();
    this.lastLogTime = this.startTime;
    this.lastLogIteration = 0;
    this.currentIter = 0;
    this.timer = null;

    this.scheduleLog();
  }
  
  public update(iter: number): void {
    this.currentIter = iter;
  }
  
  private logProgress = (): void => {
    const now = performance.now();
    const iterationsSinceLastLog = this.currentIter - this.lastLogIteration;
    const timeSinceLastLogMs = now - this.lastLogTime;
    const rate = (iterationsSinceLastLog / timeSinceLastLogMs) * this.intervalMs;
    const timeUnit = this.intervalMs >= 60000 ? 'minute' : 'second';

    logWithTime(`Iteration: ${this.currentIter + 1}/${this.total}, rate: ${rate.toLocaleString()} iterations/${timeUnit}`);

    this.lastLogTime = now;
    this.lastLogIteration = this.currentIter;

    if (this.currentIter < this.total) {
      this.scheduleLog();
    }
  }
  
  private scheduleLog(): void {
    this.timer = setTimeout(this.logProgress, this.intervalMs);
  }
  
  public finish(): void {
    if (this.timer !== null) {
      clearTimeout(this.timer);
      this.timer = null;
    }

    const totalTimeMs = performance.now() - this.startTime;
    const overallRate = (this.total / totalTimeMs) * 60000;
    logWithTime(`Completed ${this.total.toLocaleString()} iterations in ${totalTimeMs.toLocaleString()}ms (${prettyMilliseconds(totalTimeMs)}), overall rate: ${overallRate.toLocaleString()} iterations/minute`);
  }
}

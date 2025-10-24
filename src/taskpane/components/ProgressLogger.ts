import {logWithTime} from "./Common";

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
    this.startTime = Date.now();
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
    const now = Date.now();
    const iterationsSinceLastLog = this.currentIter - this.lastLogIteration;
    const timeSinceLastLog = (now - this.lastLogTime) / 1000;
    const rate = (iterationsSinceLastLog / timeSinceLastLog) * (this.intervalMs / 1000);
    const timeUnit = this.intervalMs >= 60000 ? 'minute' : 'second';
    
    logWithTime(`Iteration: ${this.currentIter + 1}/${this.total}, rate: ${rate.toFixed(2)} iterations/${timeUnit}`);
    
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
    
    const totalTime = (Date.now() - this.startTime) / 1000;
    const overallRate = (this.total / totalTime) * 60;
    logWithTime(`Completed ${this.total} iterations in ${totalTime.toFixed(2)}s, overall rate: ${overallRate.toFixed(2)} iterations/minute`);
  }
}

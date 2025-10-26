/**
 * Log with timestamp prefix
 * @param args values to log
 */
export const logWithTime = (...args: unknown[]) => {
  console.log(new Date().toISOString(), ...args);
};

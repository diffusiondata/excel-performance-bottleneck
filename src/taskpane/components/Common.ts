/**
 * Log with timestamp prefix
 */
export const logWithTime = (...args: unknown[]) => {
  console.log(new Date().toISOString(), ...args);
};

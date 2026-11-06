/** Local checks stay complete; only the explicitly selected daily CI defers
 * actual-capacity cases to the required stage acceptance. */
export const dailyCi = process.env.DOFFI_CI_MODE === 'daily';

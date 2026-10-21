export function usdToCredits(value: string): number {
  if (!/^(0|[1-9]\d{0,12})(\.\d{1,2})?$/.test(value)) throw new Error("Enter a positive USD amount with at most two decimal places.");
  const [whole, fraction = ""] = value.split(".");
  const credits = Number(BigInt(whole) * 100n + BigInt(fraction.padEnd(2, "0")));
  if (!Number.isSafeInteger(credits) || credits < 1) throw new Error("The monthly budget must fund at least one Copilot Credit.");
  return credits;
}
export function splitCredits(total: number): { standard: number; enhanced: number } {
  if (!Number.isSafeInteger(total) || total < 1) throw new Error("A positive safe whole-credit budget is required.");
  const standard = Math.floor(total / 2);
  return { standard, enhanced: total - standard };
}
export function checkAllocation(standard: number, enhanced: number, ceiling: number): void {
  if (![standard, enhanced, ceiling].every(Number.isSafeInteger) || standard < 0 || enhanced < 0 ||
      ceiling < 1 || standard > ceiling - enhanced) {
    throw new Error("Tier limits must be non-negative whole credits whose total does not exceed the approved ceiling.");
  }
}

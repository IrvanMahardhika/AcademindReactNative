const MOD = 1_000_000_000;

/**
 * Maximum income from trading one asset, starting with one share held.
 *
 * DP states per day:
 * - hold: max profit while holding a share
 * - notHold: max profit while not holding a share
 *
 * Time: O(N), Space: O(1)
 */
function solution(A: number[]): number {
  const n = A.length;
  if (n === 0) {
    return 0;
  }

  let hold = 0;
  let notHold = A[0];

  for (let i = 1; i < n; i++) {
    const price = A[i];
    const prevHold = hold;
    const prevNotHold = notHold;

    hold = Math.max(prevHold, prevNotHold - price);
    notHold = Math.max(prevNotHold, prevHold + price);
  }

  return notHold % MOD;
}

export { solution };

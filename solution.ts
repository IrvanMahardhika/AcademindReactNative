/**
 * Returns the smallest positive integer (greater than 0) not present in A.
 *
 * Only values in [1, N] can be the answer, where N = A.length.
 * Scan A once to record which of those values appear, then return the
 * first missing positive in order.
 *
 * Time: O(N), Space: O(N)
 */
function solution(A: number[]): number {
  const n = A.length;
  const present = new Set<number>();

  for (const value of A) {
    if (value > 0 && value <= n) {
      present.add(value);
    }
  }

  for (let candidate = 1; candidate <= n + 1; candidate++) {
    if (!present.has(candidate)) {
      return candidate;
    }
  }

  return n + 1;
}

export { solution };

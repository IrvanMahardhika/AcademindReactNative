/**
 * Returns the largest value X in A that appears exactly X times.
 * If no such value exists, returns 0.
 *
 * Time: O(N), Space: O(U) where U is the number of distinct values.
 */
function solution(A: number[]): number {
  const frequencies = new Map<number, number>();

  for (const value of A) {
    frequencies.set(value, (frequencies.get(value) ?? 0) + 1);
  }

  let result = 0;

  for (const [value, count] of frequencies) {
    if (value === count && value > result) {
      result = value;
    }
  }

  return result;
}

export { solution };

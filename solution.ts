/**
 * Returns the minimum number of substrings to split S into so that
 * each substring contains no repeated letters.
 *
 * Greedy: extend the current substring while all characters stay unique.
 * When the next character would repeat one already in the current substring,
 * close that substring and start a new one with that character.
 *
 * Time: O(N), Space: O(1) — at most 26 lowercase letters in the set.
 */
function solution(S: string): number {
  let segments = 1;
  const seen = new Set<string>();

  for (const char of S) {
    if (seen.has(char)) {
      segments++;
      seen.clear();
    }
    seen.add(char);
  }

  return segments;
}

export { solution };

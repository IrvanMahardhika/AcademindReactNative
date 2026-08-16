/**
 * Two inclusive windows [a, b] and [c, d] overlap iff they share at least one time point.
 */
function windowsOverlap(a, b) {
  return Math.max(a[0], b[0]) <= Math.min(a[1], b[1]);
}

/**
 * Maximum number of runners in a chain where each consecutive pair has overlapping windows.
 * Runners may be reordered. (Equivalently: longest simple path in the interval-overlap graph.)
 *
 * Uses DP over (lastRunner, visitedMask). With n runners this is O(n * 2^n), which is
 * suitable for typical interview constraints (n <= ~22). For larger n, memoization still
 * skips unreachable masks and is often faster in practice on sparse overlap graphs.
 *
 * @param {Array<[number, number]>} windows - each [start, end], inclusive
 * @returns {number}
 */
function maxRelayChain(windows) {
  const n = windows.length;
  if (n === 0) return 0;

  const adj = Array.from({ length: n }, () => []);
  for (let i = 0; i < n; i++) {
    for (let j = i + 1; j < n; j++) {
      if (windowsOverlap(windows[i], windows[j])) {
        adj[i].push(j);
        adj[j].push(i);
      }
    }
  }

  const memo = new Map();

  function dfs(last, visited) {
    const key = (BigInt(last) << BigInt(n)) | BigInt(visited);
    if (memo.has(key)) return memo.get(key);

    let best = 1;
    for (const next of adj[last]) {
      const bit = 1 << next;
      if ((visited & bit) === 0) {
        best = Math.max(best, 1 + dfs(next, visited | bit));
      }
    }

    memo.set(key, best);
    return best;
  }

  let answer = 1;
  for (let start = 0; start < n; start++) {
    answer = Math.max(answer, dfs(start, 1 << start));
  }
  return answer;
}

/**
 * Parse competitive-programming style input:
 * n
 * start_0 end_0
 * ...
 * start_{n-1} end_{n-1}
 *
 * @param {string} input
 * @returns {number}
 */
function maxRelayChainFromInput(input) {
  const lines = input
    .trim()
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const n = Number(lines[0]);
  const windows = lines.slice(1, 1 + n).map((line) => {
    const [start, end] = line.split(/\s+/).map(Number);
    return [start, end];
  });

  return maxRelayChain(windows);
}

module.exports = {
  maxRelayChain,
  maxRelayChainFromInput,
  windowsOverlap,
};

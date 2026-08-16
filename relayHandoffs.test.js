const { maxRelayChain, maxRelayChainFromInput } = require('./relayHandoffs');

function bruteForceMaxChain(windows) {
  const n = windows.length;
  const overlaps = (a, b) => Math.max(a[0], b[0]) <= Math.min(a[1], b[1]);

  const adj = Array.from({ length: n }, () => []);
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < n; j++) {
      if (i !== j && overlaps(windows[i], windows[j])) adj[i].push(j);
    }
  }

  const memo = new Map();
  function dfs(u, visited) {
    const key = `${u}:${visited}`;
    if (memo.has(key)) return memo.get(key);

    let best = 1;
    for (const v of adj[u]) {
      if ((visited & (1 << v)) === 0) {
        best = Math.max(best, 1 + dfs(v, visited | (1 << v)));
      }
    }
    memo.set(key, best);
    return best;
  }

  let ans = 0;
  for (let i = 0; i < n; i++) {
    ans = Math.max(ans, dfs(i, 1 << i));
  }
  return ans;
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

// Edge case: all disjoint -> chain of 1 runner
assertEqual(
  maxRelayChain([
    [1, 2],
    [3, 4],
    [5, 6],
  ]),
  1,
  'all disjoint'
);

// Edge case: all fully overlap -> all runners in one chain
assertEqual(
  maxRelayChain([
    [1, 10],
    [2, 9],
    [3, 8],
    [4, 7],
  ]),
  4,
  'all overlap'
);

// Classic overlapping chain
assertEqual(
  maxRelayChain([
    [1, 5],
    [3, 7],
    [6, 8],
  ]),
  3,
  'linear overlap chain'
);

// Order should not matter
assertEqual(
  maxRelayChain([
    [5, 10],
    [1, 6],
  ]),
  2,
  'reverse-friendly pair'
);

// Input parser
assertEqual(
  maxRelayChainFromInput(`3
1 5
3 7
6 8`),
  3,
  'input parser'
);

// Random small-case cross-check vs brute force
for (let trial = 0; trial < 200; trial++) {
  const n = 2 + Math.floor(Math.random() * 6);
  const windows = Array.from({ length: n }, () => {
    const a = Math.floor(Math.random() * 10);
    const b = a + Math.floor(Math.random() * 10);
    return [a, b];
  });

  const fast = maxRelayChain(windows);
  const slow = bruteForceMaxChain(windows);
  if (fast !== slow) {
    throw new Error(`Mismatch on trial ${trial}: fast=${fast}, slow=${slow}, windows=${JSON.stringify(windows)}`);
  }
}

console.log('All relayHandoffs tests passed.');

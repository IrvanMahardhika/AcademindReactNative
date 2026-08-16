/**
 * Maximum billboards with at least `minDistance` km between consecutive choices.
 * Positions must be sorted in non-decreasing order.
 *
 * Greedy: take the earliest position, then always take the next valid position
 * at or after (lastChosen + minDistance). This maximizes count on a sorted line.
 *
 * @param {number[]} positions - sorted kilometer positions
 * @param {number} minDistance - minimum required distance D
 * @returns {number}
 */
function maxBillboards(positions, minDistance) {
  const n = positions.length;
  if (n === 0) return 0;

  let count = 1;
  let last = positions[0];

  for (let i = 1; i < n; i++) {
    if (positions[i] - last >= minDistance) {
      count++;
      last = positions[i];
    }
  }

  return count;
}

/**
 * Parse stdin-style input:
 * n d
 * p1 p2 ... pn
 *
 * @param {string} input
 * @returns {number}
 */
function maxBillboardsFromInput(input) {
  const lines = input
    .trim()
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const [n, minDistance] = lines[0].split(/\s+/).map(Number);
  const positions = lines[1].split(/\s+/).map(Number).slice(0, n);

  return maxBillboards(positions, minDistance);
}

module.exports = {
  maxBillboards,
  maxBillboardsFromInput,
};

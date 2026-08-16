const { maxBillboards, maxBillboardsFromInput } = require('./billboardPlacement');

function assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

// Classic example: choose 1, 4, 7 with D=3 from [1,2,4,6,7,8]
assertEqual(maxBillboards([1, 2, 4, 6, 7, 8], 3), 3, 'sample');

// Edge: all farther than D apart -> install all
assertEqual(maxBillboards([1, 10, 25, 40], 5), 4, 'all farther apart');

// Edge: all closer than D -> only one billboard
assertEqual(maxBillboards([1, 2, 3, 4], 5), 1, 'all closer than D');

// Single position
assertEqual(maxBillboards([42], 100), 1, 'single position');

// Empty
assertEqual(maxBillboards([], 5), 0, 'empty');

// Exact spacing
assertEqual(maxBillboards([0, 5, 10, 15], 5), 4, 'exact spacing');

// Input parser
assertEqual(
  maxBillboardsFromInput(`6 3
1 2 4 6 7 8`),
  3,
  'input parser'
);

console.log('All billboardPlacement tests passed.');

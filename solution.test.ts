import { solution } from './solution';

function assertEqual(actual: number, expected: number, label: string): void {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

assertEqual(solution([1, 3, 6, 4, 1, 2]), 5, 'sample with duplicates');
assertEqual(solution([1, 2, 3]), 4, 'contiguous positives');
assertEqual(solution([-1, -3]), 1, 'no positives');
assertEqual(solution([2]), 1, 'single non-one');
assertEqual(solution([1]), 2, 'single one');

console.log('All solution tests passed.');

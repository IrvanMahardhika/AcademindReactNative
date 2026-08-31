import { solution } from './solution';

function assertEqual(actual: number, expected: number, label: string): void {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

assertEqual(solution([4, 1, 2, 3]), 6, 'example 1');
assertEqual(solution([1, 2, 3, 3, 2, 1, 5]), 7, 'example 2');
assertEqual(solution([1_000_000_000, 1, 2, 2, 1_000_000_000, 1, 1_000_000_000]), 999_999_998, 'example 3');
assertEqual(solution([5]), 5, 'single day');
assertEqual(solution([3, 3, 3]), 3, 'flat prices');

console.log('All solution tests passed.');

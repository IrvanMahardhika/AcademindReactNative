import { solution } from './solution';

function assertEqual(actual: number, expected: number, label: string): void {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

assertEqual(solution([3, 8, 2, 3, 3, 2]), 3, 'example 1');
assertEqual(solution([7, 1, 2, 8, 2]), 2, 'example 2');
assertEqual(solution([3, 1, 4, 1, 5]), 0, 'example 3');
assertEqual(solution([5, 5, 5, 5, 5]), 5, 'example 4');
assertEqual(solution([1]), 1, 'single element');

console.log('All solution tests passed.');

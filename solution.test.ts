import { solution } from './solution';

function assertEqual(actual: number, expected: number, label: string): void {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

assertEqual(solution('world'), 1, 'example 1');
assertEqual(solution('dddd'), 4, 'example 2');
assertEqual(solution('cycle'), 2, 'example 3');
assertEqual(solution('abba'), 2, 'example 4');
assertEqual(solution('abacdec'), 3, 'sample from problem statement');

console.log('All solution tests passed.');

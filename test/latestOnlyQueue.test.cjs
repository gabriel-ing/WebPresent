const test = require('node:test');
const assert = require('node:assert/strict');

const { createLatestOnlyQueue } = require('../electron/latestOnlyQueue.cjs');

async function flushMicrotasks() {
  await Promise.resolve();
  await Promise.resolve();
}

test('latest-only queue skips superseded pending values', async () => {
  const seen = [];
  const releases = [];

  const queue = createLatestOnlyQueue(async (value) => {
    seen.push(value);
    await new Promise((resolve) => {
      releases.push(resolve);
    });
  });

  const firstRun = queue(1);
  const secondRun = queue(2);
  const thirdRun = queue(3);

  await flushMicrotasks();
  assert.deepEqual(seen, [1], 'expected the first queued value to start immediately');

  releases.shift()();
  await flushMicrotasks();
  assert.deepEqual(seen, [1, 3], 'expected newer queued values to replace stale pending work');

  releases.shift()();
  await Promise.all([firstRun, secondRun, thirdRun]);
  assert.deepEqual(seen, [1, 3]);
});

test('latest-only queue continues to process later requests after becoming idle', async () => {
  const seen = [];
  const queue = createLatestOnlyQueue(async (value) => {
    seen.push(value);
  });

  await queue(1);
  await queue(2);

  assert.deepEqual(seen, [1, 2]);
});
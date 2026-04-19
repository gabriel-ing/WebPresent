function createLatestOnlyQueue(worker) {
  let activePromise = null;
  let hasPending = false;
  let pendingValue;

  function run(value) {
    activePromise = Promise.resolve(worker(value))
      .finally(() => {
        if (hasPending) {
          hasPending = false;
          return run(pendingValue);
        }

        activePromise = null;
        return undefined;
      });

    return activePromise;
  }

  return function enqueue(value) {
    if (!activePromise) {
      return run(value);
    }

    pendingValue = value;
    hasPending = true;
    return activePromise;
  };
}

module.exports = {
  createLatestOnlyQueue,
};
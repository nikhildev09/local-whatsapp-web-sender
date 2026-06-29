const defaultQueue = {
  total: 0,
  sent: 0,
  failed: 0,
  skipped: 0
};

function renderQueue(queue) {
  document.getElementById("total").textContent = queue.total || 0;
  document.getElementById("sent").textContent = queue.sent || 0;
  document.getElementById("failed").textContent = queue.failed || 0;
  document.getElementById("skipped").textContent = queue.skipped || 0;
}

chrome.storage.local.get(defaultQueue, renderQueue);

document.getElementById("clearQueue").addEventListener("click", () => {
  chrome.storage.local.set(defaultQueue, () => {
    renderQueue(defaultQueue);
  });
});

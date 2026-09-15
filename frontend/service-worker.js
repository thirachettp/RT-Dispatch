// SCDC service worker: Web Push notifications + minimal PWA installability.
// Deliberately does NOT cache app/API responses — this app relies on live
// data, so an offline-first cache would show stale task/driver state.

self.addEventListener("install", (event) => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

// Tell every open copy of the app that something changed, so it can pull fresh
// data immediately. This is what lets the page poll on a 30s timer instead of
// every few seconds: the timer is only a safety net, and a push is the real
// signal that there is something new to show.
async function notifyClientsToRefresh() {
  try {
    const clientList = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
    for (const client of clientList) {
      client.postMessage({ type: "scdc-refresh" });
    }
  } catch (e) {
    // A refresh hint failing must never stop the notification itself from
    // being shown — that notification may be the only thing telling a driver
    // there is a job waiting.
  }
}

self.addEventListener("push", (event) => {
  let data = { title: "SCDC", body: "มีการแจ้งเตือนใหม่", url: "/" };
  try {
    if (event.data) data = { ...data, ...event.data.json() };
  } catch (e) {
    if (event.data) data.body = event.data.text();
  }
  const options = {
    body: data.body,
    icon: undefined,
    badge: undefined,
    data: { url: data.url || "/" },
    tag: "scdc-notification",
    renotify: true,
  };
  event.waitUntil(
    Promise.all([
      self.registration.showNotification(data.title, options),
      notifyClientsToRefresh(),
    ])
  );
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const url = (event.notification.data && event.notification.data.url) || "/";
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((clientList) => {
      for (const client of clientList) {
        if ("focus" in client) {
          client.focus();
          if ("navigate" in client) client.navigate(url);
          // Focusing an existing window does not reload it, so ask it to pull
          // fresh data — otherwise tapping a "new task" notification could
          // land the driver on a list that does not show that task yet.
          client.postMessage({ type: "scdc-refresh" });
          return;
        }
      }
      if (self.clients.openWindow) return self.clients.openWindow(url);
    })
  );
});

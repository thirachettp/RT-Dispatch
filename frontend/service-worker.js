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
    // Served by the app itself, so it is there on every deployment.
    icon: "/icon-192.png",
    data: { url: data.url || "/" },
    // The server sends a tag per job (and one per chat thread). Every
    // notification used to share one fixed tag, which meant each new one
    // REPLACED the last: three jobs arriving together showed as one line and
    // the first two were never seen.
    tag: data.tag || ("scdc-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8)),
    renotify: true,
  };
  event.waitUntil(
    Promise.all([
      self.registration.showNotification(data.title, options),
      notifyClientsToRefresh(),
    ])
  );
});

// The browser can replace a push subscription on its own — Chrome does it
// after an update or when storage is cleared. Nothing in the page runs at that
// moment, so unless the new address is registered from here, notifications for
// this device stop dead and nothing anywhere says so: the app still shows
// "เปิดแจ้งเตือนแล้ว" while every send goes to an address that no longer exists.
self.addEventListener("pushsubscriptionchange", (event) => {
  event.waitUntil((async () => {
    try {
      const oldSub = event.oldSubscription || (await self.registration.pushManager.getSubscription());
      let key = event.oldSubscription && event.oldSubscription.options
        && event.oldSubscription.options.applicationServerKey;
      if (!key) {
        const res = await fetch("/push/vapid-public-key");
        const { publicKey } = await res.json();
        const padded = publicKey + "=".repeat((4 - publicKey.length % 4) % 4);
        const raw = atob(padded.replace(/-/g, "+").replace(/_/g, "/"));
        key = Uint8Array.from([...raw].map((ch) => ch.charCodeAt(0)));
      }
      const fresh = event.newSubscription || (await self.registration.pushManager.subscribe({
        userVisibleOnly: true,
        applicationServerKey: key,
      }));
      const json = fresh.toJSON();
      await fetch("/push/resubscribe", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          old_endpoint: oldSub ? oldSub.endpoint : "",
          endpoint: json.endpoint,
          keys: json.keys,
        }),
      });
    } catch (e) {
      // Nothing useful to do here — the page re-registers on the next sign-in.
    }
  })());
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const url = (event.notification.data && event.notification.data.url) || "/";
  // Which job this notification is about, pulled out of its own url so the
  // page gets it even when focusing an open window leaves the address bar
  // untouched — which is the usual case, and why tapping a notification used
  // to land on whatever screen happened to be open.
  let task = null, kind = "task";
  try {
    const u = new URL(url, self.location.origin);
    task = u.searchParams.get("task");
    kind = u.searchParams.get("kind") || "task";
  } catch (e) { /* a url we cannot parse carries no job */ }
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((clientList) => {
      for (const client of clientList) {
        if ("focus" in client) {
          client.focus();
          if ("navigate" in client) client.navigate(url);
          // Focusing an existing window does not reload it, so ask it to pull
          // fresh data — otherwise tapping a "new task" notification could
          // land the driver on a list that does not show that task yet.
          client.postMessage({ type: "scdc-refresh", task: task, kind: kind });
          return;
        }
      }
      if (self.clients.openWindow) return self.clients.openWindow(url);
    })
  );
});

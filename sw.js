// Minerva Daily Priorities: works offline. Bump VERSION after each update.
const VERSION="priorities-v24";
const CORE=["./","index.html","manifest.webmanifest","logo-white.png","icon.png","icon-192.png","icon-512.png"];
self.addEventListener("install",e=>{e.waitUntil(caches.open(VERSION).then(c=>Promise.all(CORE.map(u=>c.add(u).catch(()=>{})))).then(()=>self.skipWaiting()))});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==VERSION).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener("fetch",e=>{
  const req=e.request; if(req.method!=="GET")return;
  const u=new URL(req.url);
  if(u.hostname.endsWith("script.google.com")||u.hostname.endsWith("googleusercontent.com"))return; // backup calls go straight to Google
  // fresh copy when online, saved copy when offline
  e.respondWith(fetch(req,u.origin===location.origin?{cache:"no-cache"}:undefined).then(r=>{if(r.ok&&(u.origin===location.origin||/cdnjs\.cloudflare\.com|fonts\.(googleapis|gstatic)\.com/.test(u.hostname))){const c=r.clone();caches.open(VERSION).then(x=>x.put(req,c))}return r})
    .catch(()=>caches.match(req).then(h=>h||(req.mode==="navigate"?caches.match("index.html"):undefined))));
});

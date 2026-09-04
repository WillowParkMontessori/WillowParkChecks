const CACHE="willow-park-checks-v3-2-onedrive";
const CORE=["./","./index.html","./styles.css","./app.js","./manifest.webmanifest"];
const LIBS=[
  "https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js",
  "https://unpkg.com/@azure/msal-browser@5.16.0/lib/msal-browser.min.js"
];

self.addEventListener("install",event=>{
  event.waitUntil(caches.open(CACHE).then(async cache=>{
    await cache.addAll(CORE);
    for(const url of LIBS){try{await cache.add(url)}catch{}}
  }).then(()=>self.skipWaiting()));
});

self.addEventListener("activate",event=>event.waitUntil(
  caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())
));

self.addEventListener("fetch",event=>{
  if(event.request.method!=="GET") return;
  const url=new URL(event.request.url);
  const isCore=url.origin===self.location.origin;
  const isLibrary=LIBS.includes(event.request.url);
  if(!isCore && !isLibrary) return;
  event.respondWith(caches.match(event.request).then(cached=>cached||fetch(event.request).then(resp=>{
    if(resp.ok){const copy=resp.clone();caches.open(CACHE).then(c=>c.put(event.request,copy));}
    return resp;
  })));
});

const CACHE="willow-park-checks-v2-5-supabase-pdf-cloud";
const CORE=["./","./index.html","./styles.css","./app.js","./manifest.webmanifest"];
const EXTERNAL=[
 "https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js",
 "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"
];
self.addEventListener("install",event=>{
  event.waitUntil(caches.open(CACHE).then(async c=>{
    await c.addAll(CORE);
    for(const url of EXTERNAL){try{await c.add(url)}catch{}}
  }).then(()=>self.skipWaiting()));
});
self.addEventListener("activate",event=>event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener("fetch",event=>{
  if(event.request.method!=="GET")return;
  event.respondWith(caches.match(event.request).then(cached=>cached||fetch(event.request).then(resp=>{
    const copy=resp.clone();caches.open(CACHE).then(c=>c.put(event.request,copy));return resp;
  }).catch(()=>cached)));
});

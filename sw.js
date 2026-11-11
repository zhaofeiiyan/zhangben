const CACHE='feiyan-app-v1.3.1';
const BASE=new URL('./',self.location.href);
const FILES=['./','index.html','styles.css','app.js','core.js','import-tables.js','numbers-worker.js','xlsx.full.min.js','export-xlsx.js','manifest.webmanifest','manifest-v131.webmanifest','feiyan-touch-180-v131.png','feiyan-icon-192-v131.png','feiyan-icon-512-v131.png','apple-touch-icon.png','apple-touch-icon-precomposed.png','icon.svg','icon-192.png','icon-512.png','icon-192.png?v=1.2.0','icon-512.png?v=1.2.0','apple-touch-icon.png?v=1.2.0'].map(p=>new URL(p,BASE).href);
self.addEventListener('install',event=>event.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES))));
// Activate only after all pages of the previous version close, or an explicit user update request.
self.addEventListener('message',e=>{if(e.data?.type==='SKIP_WAITING')self.skipWaiting();});
self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k.startsWith('feiyan-app-')&&k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener('fetch',event=>{
 const request=event.request;
 if(request.method!=='GET')return;
 const url=new URL(request.url);
 if(url.origin!==BASE.origin||!url.pathname.startsWith(BASE.pathname))return;
 if(request.mode==='navigate'){
  event.respondWith(caches.open(CACHE).then(async cache=>{
   const cached=await cache.match(new URL('index.html',BASE).href)||await cache.match(BASE.href);
   if(cached)return cached;return fetch(request);
  }));return;
 }
 if(FILES.includes(url.href))event.respondWith(caches.open(CACHE).then(cache=>cache.match(request).then(cached=>cached||fetch(request))));
});

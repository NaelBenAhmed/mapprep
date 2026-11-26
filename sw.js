// Mise à jour automatique : ce fichier change à chaque nouvelle version (VERSION),
// le navigateur installe alors la nouvelle version et l'appli se recharge toute seule.
const VERSION='202611261010';
const C='prepa-'+VERSION;
const FILES=['./','./index.html','./entrepot.json','./manifest.webmanifest','./icon-192.png','./icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(FILES.map(f=>new Request(f,{cache:'reload'})))));self.skipWaiting()});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{const u=new URL(e.request.url);if(e.request.method!=='GET'||u.origin!==location.origin||u.pathname.endsWith('version.json')||u.pathname.endsWith('sw.js'))return;
  // Réseau d'abord (toujours la dernière version), cache si pas de réseau.
  e.respondWith(fetch(e.request,{cache:'no-cache'}).then(r=>{if(r.ok){const cp=r.clone();caches.open(C).then(c=>c.put(e.request,cp))}return r}).catch(()=>caches.match(e.request,{ignoreSearch:true}).then(r=>r||caches.match('./index.html'))))});

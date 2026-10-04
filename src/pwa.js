// Côté page : enregistrement du service worker, bouton d'installation, avis de mise à jour.
// Tous les chemins sont relatifs à la page, donc valables sous /pptanim/ comme ailleurs.

const $ = (id) => document.getElementById(id);

export function initPwa({ log = () => {} } = {}) {
  // Installation : le navigateur signale quand l'application est installable.
  let deferred = null;
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferred = e;
    $('install').hidden = false;
  });
  $('install').onclick = async () => {
    if (!deferred) return;
    deferred.prompt();
    await deferred.userChoice;
    deferred = null;
    $('install').hidden = true;
  };
  window.addEventListener('appinstalled', () => { $('install').hidden = true; log('Application installée.'); });

  if (!('serviceWorker' in navigator)) {
    $('offline').textContent = 'hors ligne : non pris en charge par ce navigateur';
    return;
  }

  // Mise à jour : la nouvelle version attend ; on propose de recharger pour l'activer.
  const offer = (reg) => {
    if (!reg.waiting || !navigator.serviceWorker.controller) return;
    $('update').hidden = false;
    $('reload').onclick = () => reg.waiting && reg.waiting.postMessage('skipWaiting');
  };
  let reloading = false;
  const hadController = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (!hadController || reloading) return; // première installation : rien à recharger
    reloading = true;
    location.reload();
  });

  navigator.serviceWorker.register('./sw.js', { scope: './' }).then((reg) => {
    offer(reg);
    reg.addEventListener('updatefound', () => {
      const sw = reg.installing;
      if (sw) sw.addEventListener('statechange', () => { if (sw.state === 'installed') offer(reg); });
    });
    return navigator.serviceWorker.ready;
  }).then(() => {
    $('offline').textContent = 'disponible hors ligne';
  }).catch((e) => {
    // En développement (dossier src/ servi tel quel), sw.js n'existe pas : c'est attendu.
    $('offline').textContent = 'hors ligne : non disponible';
    log(`Service worker non enregistré : ${e.message}`);
  });
}

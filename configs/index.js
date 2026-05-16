'use strict';

/* ============================================================
   Config registry
   ============================================================
   Plain-script (non-module) registry for StitchConfig objects. Each file in
   configs/*.js calls registerStitchConfig() at script-load time; app.js then
   resolves the active config via getActiveStitchConfig() during init.

   This file MUST load before any configs/<id>.js and before app.js.

   Phase 2 — the client code stored at localStorage['stitcher.client.v1']
   filters the registry; the chosen config id at localStorage['stitcher.config.v1']
   picks among matching configs. If either is missing or invalid the engine's
   bootstrap shows the client-code gate dialog.

   Client codes are matched case-insensitively. The reserved code 'demo'
   matches configs that opt in via clientCodes: ['demo'] for screenshots and
   marketing without exposing live client work.
*/

window.STITCH_CONFIGS = [];

window.CLIENT_KEY = 'stitcher.client.v1';
window.CONFIG_KEY = 'stitcher.config.v1';

window.registerStitchConfig = function(cfg) {
  if (!cfg || typeof cfg !== 'object') {
    console.warn('registerStitchConfig: ignoring non-object config', cfg);
    return;
  }
  if (!cfg.id) {
    console.warn('registerStitchConfig: config missing id', cfg);
    return;
  }
  if (window.STITCH_CONFIGS.some(c => c.id === cfg.id)) {
    console.warn(`registerStitchConfig: duplicate id "${cfg.id}", ignoring`);
    return;
  }
  window.STITCH_CONFIGS.push(cfg);
};

// Returns the lowercased client code from localStorage, or '' if absent.
window.getClientCode = function() {
  return (localStorage.getItem(window.CLIENT_KEY) || '').toLowerCase().trim();
};

// Configs visible to a given client code (case-insensitive). Empty array if
// no code or no matches.
window.getConfigsForClient = function(code) {
  const lc = (code || '').toLowerCase().trim();
  if (!lc) return [];
  return window.STITCH_CONFIGS.filter(c =>
    Array.isArray(c.clientCodes) && c.clientCodes.some(cc => String(cc).toLowerCase() === lc)
  );
};

// Active config: filtered by the current client code, then chosen by the
// stored config id (falling back to the first match). Returns null if the
// gate isn't satisfied — the engine treats null as "show client-code gate".
//
// Single-config short-circuit: when only one config is registered, auto-fill
// its first clientCode into localStorage on the very first load so the gate
// dialog doesn't show. Skips itself once a second config is registered —
// from then on the gate becomes meaningful again and users need to pick a
// code explicitly.
window.getActiveStitchConfig = function() {
  let code = window.getClientCode();
  if (!code && window.STITCH_CONFIGS.length === 1) {
    const only = window.STITCH_CONFIGS[0];
    const fallback = (Array.isArray(only.clientCodes) && only.clientCodes[0]) || null;
    if (fallback) {
      code = String(fallback).toLowerCase();
      localStorage.setItem(window.CLIENT_KEY, code);
    }
  }
  const available = window.getConfigsForClient(code);
  if (available.length === 0) return null;
  const storedId = localStorage.getItem(window.CONFIG_KEY);
  return available.find(c => c.id === storedId) || available[0];
};

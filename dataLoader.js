// dataLoader.js
// Local-first loader: reads /campaigns/manifest.json (or /Campaigns/manifest.json).
// Manifest supported formats:
//   { "files": ["A.xlsx","B.xlsx"] } OR ["A.xlsx","B.xlsx"] OR { "campaigns": [...] }

(function(){
  async function tryFetchJson(url){
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    return await res.json();
  }

    async function detectCampaignDir(){
    // Works for:
    // - Local Live Server (http://127.0.0.1:5500/)
    // - GitHub Pages project sites (https://user.github.io/repo/)
    // We intentionally prefer RELATIVE paths (./campaigns) so it works under a repo subpath.
    const candidates = [
      "./campaigns","campaigns","./Campaigns","Campaigns",
      // Fallbacks (rarely needed). Use with caution on GitHub Pages.
      "/campaigns","/Campaigns"
    ];
    for (const base of candidates){
      try{
        const res = await fetch(`${base}/manifest.json`, { cache:"no-store" });
        if (res.ok) return base;
      }catch(_){}
    }
    return "./campaigns";
  }

  function normalizeManifest(manifest){
    if (Array.isArray(manifest)) return manifest;
    if (manifest && Array.isArray(manifest.files)) return manifest.files;
    if (manifest && Array.isArray(manifest.campaigns)) return manifest.campaigns;
    return [];
  }

  window.DataLoader = {
    async listCampaignFiles(){
  const base = await detectCampaignDir();
  const manifest = await tryFetchJson(`${base}/manifest.json`);

  // GitHub Pages cannot auto-list directories. We rely on manifest.json,
  // but we make it tolerant:
  // - allow entries without ".xlsx" by appending it
  // - de-duplicate
  const raw = normalizeManifest(manifest)
    .filter(Boolean)
    .map(f => (typeof f==='object' && f && f.file) ? String(f.file).trim() : String(f).trim())
    .filter(Boolean);

  const normalized = raw.map(f => {
    const lower = f.toLowerCase();
    return lower.endsWith(".xlsx") ? f : (f + ".xlsx");
  });

  const seen = new Set();
  const files = [];
  for (const f of normalized){
    const key = f.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    files.push(f);
  }

  return { base, files };
},

    async loadWorkbook(base, filename){
      const url = `${base}/${encodeURIComponent(filename)}`;
      const res = await fetch(url, { cache:"no-store" });
      if (!res.ok) throw new Error(`Cannot fetch ${filename} (${res.status})`);
      const buf = await res.arrayBuffer();
      return XLSX.read(buf, { type: "array" });
    }
  };
})();

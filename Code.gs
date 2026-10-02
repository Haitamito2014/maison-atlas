/***** MAISON ATLAS — Backend (Google Apps Script) *****
 * 1) Change ADMIN_KEY ci-dessous (ton code secret vendeur).
 * 2) (Optionnel) Exécuter > setup. Sinon le tableau se prépare tout seul à la 1re visite.
 * 3) Déployer > Nouveau déploiement > Application Web
 *      Exécuter en tant que : Moi   |   Accès : Tout le monde
 * 4) Copie l'URL /exec et envoie-la à Claude.
 */
const ADMIN_KEY = 'CHANGE-MOI-123';          // <-- ton code vendeur (garde-le secret)
const CITIES = ['Casablanca', 'Mohammedia', 'Kenitra'];
const DEFAULTS = { taux: 11.515, marge: 40, prix_max: 600 };

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) { return ss_().getSheetByName(name); }

function setup() {
  const ss = ss_();
  let p = sh_('Produits') || ss.insertSheet('Produits');
  if (p.getLastRow() < 2) {
    p.clear();
    p.appendRow(['id','page','categorie','marque','nom','achat_chf','prix_barre_chf','image','actif','prix_force_mad']);
    const rows = SEED.map(r => [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], true, '']);
    p.getRange(2, 1, rows.length, 10).setValues(rows);
    p.setFrozenRows(1);
  }
  let r = sh_('Reglages') || ss.insertSheet('Reglages');
  if (r.getLastRow() < 2) {
    r.clear();
    r.appendRow(['cle','valeur']);
    Object.keys(DEFAULTS).forEach(k => r.appendRow([k, DEFAULTS[k]]));
  }
  let c = sh_('Commandes') || ss.insertSheet('Commandes');
  if (c.getLastRow() < 1) {
    c.appendRow(['n_commande','date','statut','nom','telephone','ville','adresse','note','articles','nb_pieces','total_mad','gain_mad']);
    c.setFrozenRows(1);
  }
}

function settings_() {
  const v = sh_('Reglages').getDataRange().getValues().slice(1);
  const s = Object.assign({}, DEFAULTS);
  v.forEach(r => { if (r[0] !== '') s[r[0]] = Number(r[1]); });
  return s;
}
function products_() {
  const v = sh_('Produits').getDataRange().getValues();
  const h = v.shift();
  return v.filter(r => r[0] !== '').map((r, i) => {
    const o = {}; h.forEach((k, j) => o[k] = r[j]); o._row = i + 2; return o;
  });
}
function price_(p, s) {
  const cost = Number(p.achat_chf) * s.taux;
  const sell = p.prix_force_mad !== '' && p.prix_force_mad != null && Number(p.prix_force_mad) > 0
    ? Math.round(Number(p.prix_force_mad)) : Math.round(cost * (1 + s.marge / 100));
  const old = Number(p.prix_barre_chf) * s.taux > sell ? Math.round(Number(p.prix_barre_chf) * s.taux) : 0;
  return { cost: Math.round(cost), sell: sell, old: old, gain: Math.round(sell - cost) };
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function isActive_(p) { return p.actif === true || String(p.actif).toUpperCase() === 'TRUE'; }

function catalog_() {
  const s = settings_();
  const items = products_().filter(isActive_).map(p => {
    const pr = price_(p, s);
    return { id: p.id, page: p.page, t: p.categorie, b: p.marque, n: p.nom, sell: pr.sell, old: pr.old, src: p.image };
  }).filter(x => x.page === 'promos' || x.sell <= s.prix_max);
  return { ok: true, items: items, cities: CITIES };
}

function doGet(e) {
  if (!sh_('Produits') || !sh_('Commandes') || !sh_('Reglages')) setup();
  const a = (e.parameter.action || 'catalog');
  if (a === 'catalog') return out_(catalog_());
  return out_({ ok: false, error: 'action inconnue' });
}

function doPost(e) {
  if (!sh_('Produits') || !sh_('Commandes') || !sh_('Reglages')) setup();
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'JSON invalide' }); }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (body.action === 'order') return out_(order_(body));
    if (body.key !== ADMIN_KEY) return out_({ ok: false, error: 'Code vendeur incorrect' });
    if (body.action === 'admin') return out_(admin_());
    if (body.action === 'save') return out_(save_(body));
    if (body.action === 'status') return out_(status_(body));
    return out_({ ok: false, error: 'action inconnue' });
  } finally { lock.releaseLock(); }
}

function order_(b) {
  const c = b.customer || {};
  const clean = v => String(v || '').replace(/^[=+\-@]/, "'").slice(0, 300).trim();
  if (!clean(c.nom) || String(c.telephone || '').replace(/\D/g, '').length < 9 || !clean(c.adresse)) return { ok: false, error: 'Nom, téléphone et adresse obligatoires.' };
  if (CITIES.indexOf(c.ville) < 0) return { ok: false, error: 'Livraison uniquement à Casablanca, Mohammedia et Kenitra.' };
  const s = settings_();
  const map = {}; products_().filter(isActive_).forEach(p => map[p.id] = p);
  const lines = []; let total = 0, gain = 0, n = 0;
  (b.items || []).slice(0, 50).forEach(it => {
    const p = map[it.id]; const q = Math.max(1, Math.min(10, parseInt(it.qty, 10) || 1));
    if (!p) return;
    const pr = price_(p, s);
    total += pr.sell * q; gain += pr.gain * q; n += q;
    lines.push(q + ' x ' + p.marque + ' — ' + p.nom + ' [' + p.id + '] taille ' + clean(it.size || '?') + ' — ' + pr.sell + ' MAD');
  });
  if (!lines.length) return { ok: false, error: 'Panier vide ou articles indisponibles.' };
  const sh = sh_('Commandes');
  const num = 'MA-' + Utilities.formatDate(new Date(), 'Africa/Casablanca', 'yyMMdd') + '-' + (sh.getLastRow());
  const tel = String(c.telephone || '').replace(/[^\d+ ]/g, '').slice(0, 20);
  sh.appendRow([num, new Date(), 'Nouvelle', clean(c.nom), "'" + tel, c.ville, clean(c.adresse), clean(c.note), lines.join('\n'), n, total, gain]);
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      'Nouvelle commande ' + num + ' — ' + total + ' MAD (' + c.ville + ')',
      'Client : ' + clean(c.nom) + '\nTél : ' + String(c.telephone || '').replace(/[^\d+ ]/g, '') + '\nVille : ' + c.ville + '\nAdresse : ' + clean(c.adresse) +
      '\nNote : ' + clean(c.note) + '\n\n' + lines.join('\n') + '\n\nTotal à encaisser (espèces) : ' + total + ' MAD\nGain : ' + gain + ' MAD');
  } catch (err) {}
  return { ok: true, num: num, total: total };
}

function admin_() {
  const s = settings_();
  const prods = products_().map(p => {
    const pr = price_(p, s);
    return { id: p.id, page: p.page, t: p.categorie, b: p.marque, n: p.nom, achat: Number(p.achat_chf), rrp: Number(p.prix_barre_chf),
      src: p.image, actif: isActive_(p), force: p.prix_force_mad === '' ? '' : Number(p.prix_force_mad), cost: pr.cost, sell: pr.sell, gain: pr.gain };
  });
  const ov = sh_('Commandes').getDataRange().getValues(); const oh = ov.shift();
  const orders = ov.filter(r => r[0] !== '').map((r, i) => { const o = { _row: i + 2 }; oh.forEach((k, j) => o[k] = r[j] instanceof Date ? r[j].toISOString() : r[j]); return o; }).reverse();
  return { ok: true, settings: s, products: prods, orders: orders };
}

function save_(b) {
  if (b.settings) {
    const r = sh_('Reglages'); const v = r.getDataRange().getValues();
    Object.keys(DEFAULTS).forEach(k => {
      if (b.settings[k] === undefined || isNaN(Number(b.settings[k]))) return;
      const i = v.findIndex(x => x[0] === k);
      if (i > 0) r.getRange(i + 1, 2).setValue(Number(b.settings[k])); else r.appendRow([k, Number(b.settings[k])]);
    });
  }
  if (b.products && b.products.length) {
    const sh = sh_('Produits'); const rows = {}; products_().forEach(p => rows[p.id] = p._row);
    b.products.forEach(u => {
      const row = rows[u.id]; if (!row) return;
      if (u.actif !== undefined) sh.getRange(row, 9).setValue(!!u.actif);
      if (u.force !== undefined) sh.getRange(row, 10).setValue(u.force === '' || u.force === null ? '' : Number(u.force));
      if (u.achat !== undefined && !isNaN(Number(u.achat))) sh.getRange(row, 6).setValue(Number(u.achat));
    });
  }
  return admin_();
}

function status_(b) {
  const ok = ['Nouvelle','Confirmée','En préparation','Livrée','Annulée'];
  if (ok.indexOf(b.statut) < 0) return { ok: false, error: 'statut invalide' };
  const sh = sh_('Commandes'); const v = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues();
  const i = v.findIndex(r => r[0] === b.num); if (i < 0) return { ok: false, error: 'commande introuvable' };
  sh.getRange(i + 2, 3).setValue(b.statut);
  return { ok: true };
}

// Données de départ (privées : prix d'achat en CHF)
const SEED = [["C001","chemises","chemise","Pierre Cardin","Chemise manches longues – modern fit",17.0,66.0,"images/chemise_01.jpg"],["C002","chemises","chemise","Venti","Chemise manches longues – modern fit",19.0,47.0,"images/chemise_02.jpg"],["C003","chemises","chemise","Jack & Jones","Chemise manches longues Charge",20.0,33.0,"images/chemise_03.jpg"],["C004","chemises","chemise","Sublevel","Chemise manches longues – regular fit",20.0,47.0,"images/chemise_04.jpg"],["C005","chemises","chemise","Hatico","Chemise manches longues – modern fit",21.0,57.0,"images/chemise_05.jpg"],["C006","chemises","chemise","Pure","Chemise manches courtes – modern fit",21.0,57.0,"images/chemise_06.jpg"],["C007","chemises","chemise","Dan John","Chemise manches longues – slim fit",22.0,38.0,"images/chemise_07.jpg"],["C008","chemises","chemise","Matinique","Chemise manches longues Matrostol BD",22.0,66.0,"images/chemise_08.jpg"],["C009","chemises","chemise","Tom Tailor","Chemise manches longues – relaxed fit",22.0,47.0,"images/chemise_09.jpg"],["C010","chemises","chemise","Eterna","Chemise manches longues – casual fit",24.0,66.0,"images/chemise_10.jpg"],["C011","chemises","chemise","Jack & Jones","Chemise en lin mélangé Shelo – regular fit",24.0,43.0,"images/chemise_11.jpg"],["C012","chemises","chemise","Olymp","Chemise manches longues – super slim fit",24.0,75.0,"images/chemise_12.jpg"],["C013","chemises","chemise","S.Oliver","Chemise manches longues – regular fit",24.0,57.0,"images/chemise_13.jpg"],["C014","chemises","chemise","Seidensticker","Chemise manches longues – regular fit",24.0,57.0,"images/chemise_14.jpg"],["C015","chemises","chemise","Selected Homme","Chemise manches longues – slim fit",24.0,47.0,"images/chemise_15.jpg"],["C016","chemises","chemise","Camicissima","Chemise manches longues – slim fit",26.0,66.0,"images/chemise_16.jpg"],["C017","chemises","chemise","Edwin","Chemise manches longues Labour",26.0,121.0,"images/chemise_17.jpg"],["C018","chemises","chemise","Pierre Cardin","Chemise manches longues – casual fit",26.0,66.0,"images/chemise_18.jpg"],["C019","chemises","chemise","Ben Sherman","Chemise manches longues",27.0,70.0,"images/chemise_19.jpg"],["C020","chemises","chemise","Dan John","Chemise Fantasia Foglia – slim fit",27.0,38.0,"images/chemise_20.jpg"],["C021","chemises","chemise","Jack & Jones","Chemise Blaactive – slim fit",27.0,47.0,"images/chemise_21.jpg"],["C022","chemises","chemise","Seidensticker","Chemise manches courtes – body fit",27.0,57.0,"images/chemise_22.jpg"],["C023","chemises","chemise","Selected Homme","Chemise en lin mélangé Claus",27.0,57.0,"images/chemise_23.jpg"],["C024","chemises","chemise","Ben Sherman","Chemise manches longues",28.0,80.0,"images/chemise_24.jpg"],["C025","chemises","chemise","Gianni Feraud","Chemise manches longues – slim fit",28.0,94.0,"images/chemise_25.jpg"],["C026","chemises","chemise","Lawrence Grey","Chemise manches longues",28.0,75.0,"images/chemise_26.jpg"],["C027","chemises","chemise","Remus Uomo","Chemise Frank – slim fit",28.0,75.0,"images/chemise_27.jpg"],["C028","chemises","chemise","Selected Homme","Chemise manches longues – regular fit",28.0,57.0,"images/chemise_28.jpg"],["C029","chemises","chemise","Crew Clothing Company","Chemise manches longues – classic fit",29.0,64.0,"images/chemise_29.jpg"],["C030","chemises","chemise","Eterna","Chemise manches longues – slim fit",29.0,66.0,"images/chemise_30.jpg"],["C031","chemises","chemise","Hechter Paris","Chemise manches longues – regular fit",29.0,84.0,"images/chemise_31.jpg"],["C032","chemises","chemise","Jack & Jones","Surchemise Travis",29.0,47.0,"images/chemise_32.jpg"],["C033","chemises","chemise","Pure","Chemise technique col requin",29.0,112.0,"images/chemise_33.jpg"],["C034","chemises","chemise","Seidensticker","Chemise manches longues",29.0,57.0,"images/chemise_34.jpg"],["C035","chemises","chemise","Ben Sherman","Surchemise",32.0,93.0,"images/chemise_35.jpg"],["C036","chemises","chemise","Claudio Campione","Chemise manches longues – modern fit",32.0,84.0,"images/chemise_36.jpg"],["C037","chemises","chemise","Fil Noir","Chemise manches longues",32.0,120.0,"images/chemise_37.jpg"],["C038","chemises","chemise","Jack & Jones","Surchemise",32.0,69.9,"images/chemise_38.jpg"],["C039","chemises","chemise","Olymp","Chemise manches longues – comfort fit",32.0,57.0,"images/chemise_39.jpg"],["C040","chemises","chemise","Selected Homme","Chemise Mance – slim fit",32.0,57.0,"images/chemise_40.jpg"],["L001","luxe","chemise","Michael Kors","Chemise manches longues – slim fit",35,103,"images-luxe2/lx_001.jpg"],["L002","luxe","chemise","Napapijri","Chemise manches longues",35,84,"images-luxe2/lx_002.jpg"],["L003","luxe","chemise","Hugo","Chemise manches courtes",36,66,"images-luxe2/lx_003.jpg"],["L004","luxe","tshirt","Iceberg","T-shirt Vertical",13,61,"images-luxe2/lx_004.jpg"],["L005","luxe","tshirt","Iceberg","T-shirt Vertical",14,61,"images-luxe2/lx_005.jpg"],["L006","luxe","tshirt","Plein Sport","T-shirt col rond",14,70,"images-luxe2/lx_006.jpg"],["L007","luxe","tshirt","Iceberg","T-shirt Vertical",16,61,"images-luxe2/lx_007.jpg"],["L008","luxe","tshirt","Plein Sport","T-shirt col rond",17,70,"images-luxe2/lx_008.jpg"],["L009","luxe","tshirt","Plein Sport","Polo",18,74,"images-luxe2/lx_009.jpg"],["L010","luxe","tshirt","Plein Sport","T-shirt col rond",19,80,"images-luxe2/lx_010.jpg"],["L011","luxe","tshirt","Boss Green","T-shirt",20,57,"images-luxe2/lx_011.jpg"],["L012","luxe","tshirt","Plein Sport","T-shirt col rond",20,70,"images-luxe2/lx_012.jpg"],["L013","luxe","tshirt","Plein Sport","T-shirt col rond",20,70,"images-luxe2/lx_013.jpg"],["L014","luxe","tshirt","Boss Green","T-shirt",21,57,"images-luxe2/lx_014.jpg"],["L015","luxe","tshirt","Plein Sport","Polo",21,74,"images-luxe2/lx_015.jpg"],["L016","luxe","tshirt","Plein Sport","T-shirt col rond",21,70,"images-luxe2/lx_016.jpg"],["L017","luxe","tshirt","Plein Sport","T-shirt col rond",21,74,"images-luxe2/lx_017.jpg"],["L018","luxe","tshirt","Plein Sport","T-shirt col rond",21,70,"images-luxe2/lx_018.jpg"],["L019","luxe","tshirt","Plein Sport","T-shirt",21,57,"images-luxe2/lx_019.jpg"],["L020","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",21,38,"images-luxe2/lx_020.jpg"],["L021","luxe","tshirt","Plein Sport","T-shirt col rond",22,80,"images-luxe2/lx_021.jpg"],["L022","luxe","tshirt","Plein Sport","T-shirt col rond",22,74,"images-luxe2/lx_022.jpg"],["L023","luxe","tshirt","Plein Sport","T-shirt col rond",22,70,"images-luxe2/lx_023.jpg"],["L024","luxe","tshirt","Plein Sport","T-shirt col rond",22,70,"images-luxe2/lx_024.jpg"],["L025","luxe","tshirt","Plein Sport","T-shirt",22,57,"images-luxe2/lx_025.jpg"],["L026","luxe","tshirt","Plein Sport","T-shirt",22,61,"images-luxe2/lx_026.jpg"],["L027","luxe","tshirt","Baldessarini","T-shirt col rond",24,56,"images-luxe2/lx_027.jpg"],["L028","luxe","tshirt","Calvin Klein","T-shirt",24,38,"images-luxe2/lx_028.jpg"],["L029","luxe","tshirt","Plein Sport","T-shirt col rond",24,70,"images-luxe2/lx_029.jpg"],["L030","luxe","tshirt","Plein Sport","T-shirt col rond",24,74,"images-luxe2/lx_030.jpg"],["L031","luxe","tshirt","Tommy Hilfiger","T-shirt Essential",24,44.9,"images-luxe2/lx_031.jpg"],["L032","luxe","tshirt","Tommy Hilfiger","T-shirt",24,44.9,"images-luxe2/lx_032.jpg"],["L033","luxe","tshirt","Baldessarini","T-shirt col rond",27,56,"images-luxe2/lx_033.jpg"],["L034","luxe","tshirt","Baldessarini","T-shirt col rond",27,56,"images-luxe2/lx_034.jpg"],["L035","luxe","tshirt","Napapijri","T-shirt manches longues",27,44,"images-luxe2/lx_035.jpg"],["L036","luxe","tshirt","Napapijri","Polo – regular fit",27,71,"images-luxe2/lx_036.jpg"],["L037","luxe","tshirt","Napapijri","Polo – regular fit",27,71,"images-luxe2/lx_037.jpg"],["L038","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",27,44.9,"images-luxe2/lx_038.jpg"],["L039","luxe","tshirt","EA7 Emporio Armani","T-shirt de sport",28,75,"images-luxe2/lx_039.jpg"],["L040","luxe","tshirt","Napapijri","T-shirt manches longues",28,44,"images-luxe2/lx_040.jpg"],["L041","luxe","tshirt","Napapijri","Polo – regular fit",28,71,"images-luxe2/lx_041.jpg"],["L042","luxe","tshirt","Plein Sport","Polo",28,80,"images-luxe2/lx_042.jpg"],["L043","luxe","tshirt","Plein Sport","T-shirt",28,61,"images-luxe2/lx_043.jpg"],["L044","luxe","tshirt","Guess","T-shirt col rond",29,45,"images-luxe2/lx_044.jpg"],["L045","luxe","tshirt","Napapijri","Polo – regular fit",29,72,"images-luxe2/lx_045.jpg"],["L046","luxe","tshirt","Plein Sport","Polo",29,80,"images-luxe2/lx_046.jpg"],["L047","luxe","tshirt","Plein Sport","Polo",29,80,"images-luxe2/lx_047.jpg"],["L048","luxe","tshirt","The Kooples","T-shirt manches longues",29,117,"images-luxe2/lx_048.jpg"],["L049","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_049.jpg"],["L050","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_050.jpg"],["L051","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_051.jpg"],["L052","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_052.jpg"],["L053","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_053.jpg"],["L054","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_054.jpg"],["L055","luxe","tshirt","Tommy Hilfiger","T-shirt Essential",29,38,"images-luxe2/lx_055.jpg"],["L056","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_056.jpg"],["L057","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_057.jpg"],["L058","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_058.jpg"],["L059","luxe","tshirt","Tommy Hilfiger","T-shirt Essential",29,44.9,"images-luxe2/lx_059.jpg"],["L060","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_060.jpg"],["L061","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_061.jpg"],["L062","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_062.jpg"],["L063","luxe","tshirt","Tommy Hilfiger","T-shirt Essential",29,44.9,"images-luxe2/lx_063.jpg"],["L064","luxe","tshirt","Tommy Hilfiger","T-shirt",29,38,"images-luxe2/lx_064.jpg"],["L065","luxe","tshirt","Tommy Hilfiger","T-shirt",29,44.9,"images-luxe2/lx_065.jpg"],["L066","luxe","tshirt","Tommy Hilfiger","T-shirt",29,44.9,"images-luxe2/lx_066.jpg"],["L067","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,47,"images-luxe2/lx_067.jpg"],["L068","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_068.jpg"],["L069","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_069.jpg"],["L070","luxe","tshirt","Tommy Hilfiger","T-shirt Essential",29,38,"images-luxe2/lx_070.jpg"],["L071","luxe","tshirt","Tommy Hilfiger","T-shirt Essential",29,44.9,"images-luxe2/lx_071.jpg"],["L072","luxe","tshirt","Tommy Hilfiger","T-shirt",29,47,"images-luxe2/lx_072.jpg"],["L073","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",29,38,"images-luxe2/lx_073.jpg"],["L074","luxe","tshirt","Hackett London","T-shirt manches longues",32,56,"images-luxe2/lx_074.jpg"],["L075","luxe","tshirt","Napapijri","T-shirt manches longues",32,47,"images-luxe2/lx_075.jpg"],["L076","luxe","tshirt","Napapijri","T-shirt manches longues",32,44,"images-luxe2/lx_076.jpg"],["L077","luxe","tshirt","Napapijri","Polo – regular fit",32,71,"images-luxe2/lx_077.jpg"],["L078","luxe","tshirt","Plein Sport","Polo",32,80,"images-luxe2/lx_078.jpg"],["L079","luxe","tshirt","Tommy Hilfiger","T-shirt manches longues",32,54.9,"images-luxe2/lx_079.jpg"],["L080","luxe","tshirt","Tommy Hilfiger","T-shirt manches longues",32,59.9,"images-luxe2/lx_080.jpg"],["L081","luxe","tshirt","Tommy Hilfiger","T-shirt",32,47,"images-luxe2/lx_081.jpg"],["L082","luxe","tshirt","Tommy Hilfiger","T-shirt",32,59.9,"images-luxe2/lx_082.jpg"],["L083","luxe","tshirt","EA7 Emporio Armani","T-shirt manches longues",33,66,"images-luxe2/lx_083.jpg"],["L084","luxe","tshirt","Guess","T-shirt manches longues",33,43,"images-luxe2/lx_085.jpg"],["L085","luxe","tshirt","Joop!","Polo piqué",33,57,"images-luxe2/lx_086.jpg"],["L086","luxe","tshirt","Lacoste","Lot de 3 T-shirts",33,43,"images-luxe2/lx_087.jpg"],["L087","luxe","tshirt","Napapijri","Polo manches longues",33,75,"images-luxe2/lx_088.jpg"],["L088","luxe","tshirt","Napapijri","Polo – regular fit Emira",33,82,"images-luxe2/lx_089.jpg"],["L089","luxe","tshirt","Napapijri","Polo manches longues",33,75,"images-luxe2/lx_090.jpg"],["L090","luxe","tshirt","Tommy Hilfiger","T-shirt manches longues",33,54.9,"images-luxe2/lx_091.jpg"],["L091","luxe","tshirt","Tommy Hilfiger","T-shirt manches longues",33,59.9,"images-luxe2/lx_092.jpg"],["L092","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",33,47,"images-luxe2/lx_093.jpg"],["L093","luxe","tshirt","Boss Orange","T-shirt Tegood",35,47,"images-luxe2/lx_094.jpg"],["L094","luxe","tshirt","Diesel","T-shirt manches longues T-Just-Lis-Div",35,70,"images-luxe2/lx_095.jpg"],["L095","luxe","tshirt","EA7 Emporio Armani","T-shirt de sport",35,80,"images-luxe2/lx_096.jpg"],["L096","luxe","tshirt","Hugo","Lot de 2 T-shirts Blue 2",35,59,"images-luxe2/lx_097.jpg"],["L097","luxe","tshirt","Plein Sport","T-shirt",35,57,"images-luxe2/lx_098.jpg"],["L098","luxe","tshirt","Boss Orange","T-shirt Tegood",36,47,"images-luxe2/lx_099.jpg"],["L099","luxe","tshirt","Boss Orange","T-shirt Tegood",36,47,"images-luxe2/lx_100.jpg"],["L100","luxe","tshirt","Hugo","Lot de 2 T-shirts Blue 2",36,59,"images-luxe2/lx_101.jpg"],["L101","luxe","tshirt","Karl Lagerfeld","T-shirt manches longues",36,94,"images-luxe2/lx_102.jpg"],["L102","luxe","tshirt","Lacoste","T-shirt col rond",36,66,"images-luxe2/lx_103.jpg"],["L103","luxe","tshirt","Lacoste","T-shirt col rond",36,57,"images-luxe2/lx_104.jpg"],["L104","luxe","tshirt","Lacoste","T-shirt col rond",36,66,"images-luxe2/lx_105.jpg"],["L105","luxe","tshirt","Lacoste","T-shirt col rond",36,61,"images-luxe2/lx_106.jpg"],["L106","luxe","tshirt","Napapijri","T-shirt manches longues",36,47,"images-luxe2/lx_107.jpg"],["L107","luxe","tshirt","Napapijri","Polo manches longues",36,90,"images-luxe2/lx_108.jpg"],["L108","luxe","tshirt","Plein Sport","T-shirt",36,57,"images-luxe2/lx_109.jpg"],["L109","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",36,47,"images-luxe2/lx_110.jpg"],["L110","luxe","tshirt","Tommy Hilfiger","T-shirt col rond Script",36,47,"images-luxe2/lx_111.jpg"],["L111","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",36,47,"images-luxe2/lx_112.jpg"],["L112","luxe","tshirt","Tommy Hilfiger","T-shirt Performance",36,56,"images-luxe2/lx_113.jpg"],["L113","luxe","tshirt","Tommy Hilfiger","T-shirt col rond",36,59.9,"images-luxe2/lx_114.jpg"],["L114","luxe","tshirt","Tommy Hilfiger","T-shirt col rond Script",36,47,"images-luxe2/lx_115.jpg"],["L115","luxe","tshirt","Tommy Hilfiger","T-shirt manches longues",36,59.9,"images-luxe2/lx_116.jpg"],["L116","luxe","tshirt","Tommy Hilfiger","T-shirt",36,47,"images-luxe2/lx_117.jpg"],["L117","luxe","tshirt","Tommy Hilfiger","T-shirt",36,47,"images-luxe2/lx_118.jpg"],["P001","promos","veste","Thomas Goodwin","Parka matelassée",35,121,"images-promos/promo_01.jpg"],["P002","promos","veste","S4 Jackets","Manteau mi-saison",38,149,"images-promos/promo_02.jpg"],["P003","promos","veste","S4 Jackets","Manteau matelassé",47,205,"images-promos/promo_03.jpg"],["P004","promos","veste","Benvenuto.","Manteau mi-saison",50,214,"images-promos/promo_04.jpg"],["P005","promos","veste","S4 Jackets","Manteau matelassé",47,205,"images-promos/promo_05.jpg"],["P006","promos","veste","Kensington Eastside","Manteau mi-saison",52,93,"images-promos/promo_06.jpg"],["P007","promos","veste","Kensington Eastside","Manteau mi-saison",52,93,"images-promos/promo_07.jpg"],["P008","promos","veste","S.Oliver Black Label","Manteau en laine mélangée",52,186,"images-promos/promo_08.jpg"],["P009","promos","veste","S4 Jackets","Veste matelassée",24,94,"images-promos/promo_09.jpg"],["P010","promos","veste","S4 Jackets","Veste matelassée",26,94,"images-promos/promo_10.jpg"],["P011","promos","veste","Sublevel","Veste mi-saison",24,47,"images-promos/promo_11.jpg"],["P012","promos","chemise","Sublevel","Surchemise (shacket)",27,57,"images-promos/promo_12.jpg"],["P013","promos","chemise","Sublevel","Surchemise (shacket)",27,57,"images-promos/promo_13.jpg"],["P014","promos","chemise","Jack & Jones","Surchemise",29,47,"images-promos/promo_14.jpg"],["P015","promos","chemise","Jack & Jones","Surchemise",29,47,"images-promos/promo_15.jpg"],["P016","promos","veste","Brave Soul","Veste mi-saison",29,59,"images-promos/promo_16.jpg"],["P017","promos","veste","Only & Sons","Blouson bomber",28,57,"images-promos/promo_17.jpg"],["P018","promos","tshirt","Eight2Nine","T-shirt manches longues",12,24,"images-promos/promo_18.jpg"],["P019","promos","tshirt","Eight2Nine","T-shirt manches longues",12,24,"images-promos/promo_19.jpg"],["P020","promos","tshirt","Iceberg","T-shirt",13,61,"images-promos/promo_20.jpg"],["P021","promos","tshirt","Sublevel","T-shirt manches longues",13,33,"images-promos/promo_21.jpg"],["P022","promos","tshirt","Puma","T-shirt col rond",15,22,"images-promos/promo_22.jpg"],["P023","promos","tshirt","Plein Sport","Polo",18,74,"images-promos/promo_23.jpg"],["P024","promos","tshirt","Joma","Polo",21,52,"images-promos/promo_24.jpg"],["P025","promos","tshirt","Nike","Polo",20,27,"images-promos/promo_25.jpg"],["P026","promos","tshirt","U.S. Polo Assn.","Polo – regular fit",21,93,"images-promos/promo_26.jpg"],["P027","promos","tshirt","U.S. Polo Assn.","Polo – regular fit",21,93,"images-promos/promo_27.jpg"],["P028","promos","tshirt","Puma","Polo",24,50,"images-promos/promo_28.jpg"],["P029","promos","tshirt","Nike","Polo piqué",24,33,"images-promos/promo_29.jpg"],["P030","promos","tshirt","U.S. Polo Assn.","Polo – regular fit",21,93,"images-promos/promo_30.jpg"],["P031","promos","tshirt","Napapijri","Polo – regular fit",27,71,"images-promos/promo_31.jpg"],["P032","promos","tshirt","Napapijri","Polo – regular fit",27,71,"images-promos/promo_32.jpg"],["P033","promos","tshirt","Ellesse","Polo",28,57,"images-promos/promo_33.jpg"],["P034","promos","tshirt","Plein Sport","Polo",28,80,"images-promos/promo_34.jpg"],["P035","promos","chaussures","Puma","Baskets",24,47,"images-promos/promo_35.jpg"],["P036","promos","chaussures","Bugatti","Baskets",28,47,"images-promos/promo_36.jpg"],["P037","promos","chaussures","U.S. Polo Assn.","Baskets",29,93,"images-promos/promo_37.jpg"],["P038","promos","chaussures","Pepe Jeans","Baskets",29,64.9,"images-promos/promo_38.jpg"],["P039","promos","chaussures","Crocs","Sabots",32,52,"images-promos/promo_39.jpg"],["P040","promos","chaussures","Crocs","Sabots",32,52,"images-promos/promo_40.jpg"],["P041","promos","chaussures","Crocs","Sabots",32,57,"images-promos/promo_41.jpg"],["P042","promos","chaussures","Ellesse","Baskets",32,47,"images-promos/promo_42.jpg"]];

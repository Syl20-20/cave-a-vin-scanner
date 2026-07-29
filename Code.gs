/**
 * ============================================================
 * CAVE À VIN - Google Apps Script Backend
 * ============================================================
 * Structure Google Sheets attendue :
 *
 * Feuille "Inventaire" (colonnes, ligne 1 = en-têtes) :
 * A: ID              (string, uuid)
 * B: CodeBarre       (string)
 * C: Nom             (string)
 * D: Producteur      (string)
 * E: Type            (string: Rouge/Blanc/Rosé/Mousseux/Fortifié/Autre)
 * F: Pays            (string)
 * G: Region          (string)
 * H: Cepage          (string)
 * I: Millesime       (number ou "N.M.")
 * J: Format           (string, ex "750 ml")
 * K: PrixSAQ         (number)
 * L: ImageURL        (string, url)
 * M: PastilleGout    (string)
 * N: QteCellier      (number)
 * O: QteLocker       (number)
 * P: QteFrigo        (number)
 * Q: Note            (number, /10)
 * R: Commentaire     (string)
 * S: DateAjout       (date ISO string)
 * T: DateMAJ         (date ISO string)
 *
 * Feuille "Historique" (colonnes, ligne 1 = en-têtes) :
 * A: ID              (string, uuid)
 * B: DateHeure       (date ISO string)
 * C: CodeBarre       (string)
 * D: Nom             (string)
 * E: Action          (string: Ajout/Consommation/Deplacement/Modification/Suppression)
 * F: EmplacementSource      (string)
 * G: EmplacementDestination (string)
 * H: QuantiteAvant   (number)
 * I: QuantiteApres   (number)
 * J: Commentaire     (string)
 * ============================================================
 */

// ---- CONFIGURATION ----
// Remplacez par l'ID de votre Google Sheet (dans l'URL entre /d/ et /edit)
const SHEET_ID = '14Y4Btdv8fhEnwf7cRI6diPqLQ5emQ4pajaEO_Ag-x_s';
const SHEET_INVENTAIRE = 'Inventaire';
const SHEET_HISTORIQUE = 'Historique';

const INVENTAIRE_HEADERS = [
  'ID','CodeBarre','Nom','Producteur','Type','Pays','Region','Cepage',
  'Millesime','Format','PrixSAQ','ImageURL','PastilleGout',
  'QteCellier','QteLocker','QteFrigo','Note','Commentaire','DateAjout','DateMAJ'
];

const HISTORIQUE_HEADERS = [
  'ID','DateHeure','CodeBarre','Nom','Action',
  'EmplacementSource','EmplacementDestination','QuantiteAvant','QuantiteApres','Commentaire'
];

// ============================================================
// SERVIR LA WEB APP
// ============================================================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  // Récupère un éventuel code-barres passé en paramètre d'URL
  // (utilisé par la page scanner.html hébergée sur GitHub Pages, hors iframe Apps Script)
  template.barcodeParam = (e && e.parameter && e.parameter.barcode) ? e.parameter.barcode : '';

  return template.evaluate()
    .setTitle('Ma Cave à Vin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Permet d'inclure des fichiers HTML/JS/CSS partiels si besoin (non utilisé par défaut)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// UTILITAIRES SHEET
// ============================================================
function getSheet_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = name === SHEET_INVENTAIRE ? INVENTAIRE_HEADERS : HISTORIQUE_HEADERS;
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows
    .map((row, i) => {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = row[idx]);
      obj._row = i + 2; // ligne réelle dans le sheet (1-based + header)
      return obj;
    })
    .filter(obj => obj.ID); // ignore lignes vides
}

function generateId_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}

// ============================================================
// CRUD - INVENTAIRE
// ============================================================

/**
 * Retourne tout l'inventaire (tableau d'objets JS)
 */
function getInventaire() {
  const sheet = getSheet_(SHEET_INVENTAIRE);
  return sheetToObjects_(sheet);
}

/**
 * Cherche une bouteille par code-barres. Retourne null si absente.
 */
function findByBarcode(barcode) {
  const items = getInventaire();
  const found = items.find(it => String(it.CodeBarre) === String(barcode));
  return found || null;
}

/**
 * Ajoute une nouvelle bouteille (objet "wine" venant du front-end)
 * wine = { CodeBarre, Nom, Producteur, Type, Pays, Region, Cepage,
 *          Millesime, Format, PrixSAQ, ImageURL, PastilleGout,
 *          QteCellier, QteLocker, QteFrigo, Note, Commentaire }
 */
function addWine(wine) {
  const sheet = getSheet_(SHEET_INVENTAIRE);
  const id = generateId_();
  const now = nowIso_();

  const row = [
    id,
    wine.CodeBarre || '',
    wine.Nom || '',
    wine.Producteur || '',
    wine.Type || '',
    wine.Pays || '',
    wine.Region || '',
    wine.Cepage || '',
    wine.Millesime || '',
    wine.Format || '750 ml',
    wine.PrixSAQ || '',
    wine.ImageURL || '',
    wine.PastilleGout || '',
    Number(wine.QteCellier) || 0,
    Number(wine.QteLocker) || 0,
    Number(wine.QteFrigo) || 0,
    wine.Note || '',
    wine.Commentaire || '',
    now,
    now
  ];
  sheet.appendRow(row);

  logHistorique_({
    CodeBarre: wine.CodeBarre,
    Nom: wine.Nom,
    Action: 'Ajout',
    EmplacementSource: '',
    EmplacementDestination: 'Multiple',
    QuantiteAvant: 0,
    QuantiteApres: (Number(wine.QteCellier)||0) + (Number(wine.QteLocker)||0) + (Number(wine.QteFrigo)||0),
    Commentaire: 'Nouvelle bouteille ajoutée à l\'inventaire'
  });

  return { success: true, id: id };
}

/**
 * Met à jour une bouteille existante par ID.
 * updates = objet partiel avec les champs à modifier
 */
function updateWine(id, updates) {
  const sheet = getSheet_(SHEET_INVENTAIRE);
  const items = sheetToObjects_(sheet);
  const item = items.find(it => it.ID === id);
  if (!item) return { success: false, error: 'Bouteille introuvable' };

  const headers = INVENTAIRE_HEADERS;
  const rowIndex = item._row;
  const currentRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  headers.forEach((h, idx) => {
    if (h === 'DateAjout') return; // ne jamais écraser la date d'ajout
    if (updates.hasOwnProperty(h)) {
      currentRow[idx] = updates[h];
    }
  });
  // toujours mettre à jour DateMAJ
  currentRow[headers.indexOf('DateMAJ')] = nowIso_();

  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([currentRow]);
  return { success: true };
}

/**
 * Supprime une bouteille par ID (suppression définitive de la ligne)
 */
function deleteWine(id) {
  const sheet = getSheet_(SHEET_INVENTAIRE);
  const items = sheetToObjects_(sheet);
  const item = items.find(it => it.ID === id);
  if (!item) return { success: false, error: 'Bouteille introuvable' };

  logHistorique_({
    CodeBarre: item.CodeBarre,
    Nom: item.Nom,
    Action: 'Suppression',
    EmplacementSource: 'Toutes',
    EmplacementDestination: '',
    QuantiteAvant: (Number(item.QteCellier)||0)+(Number(item.QteLocker)||0)+(Number(item.QteFrigo)||0),
    QuantiteApres: 0,
    Commentaire: 'Bouteille supprimée de l\'inventaire'
  });

  sheet.deleteRow(item._row);
  return { success: true };
}

/**
 * Ajuste la quantité d'un emplacement donné (+1 / -1 / valeur libre)
 * emplacement = 'QteCellier' | 'QteLocker' | 'QteFrigo'
 * delta = nombre (peut être négatif)
 * action = 'Consommation' | 'Ajout' | 'Ajustement'
 */
function adjustQuantity(id, emplacement, delta, action) {
  const sheet = getSheet_(SHEET_INVENTAIRE);
  const items = sheetToObjects_(sheet);
  const item = items.find(it => it.ID === id);
  if (!item) return { success: false, error: 'Bouteille introuvable' };

  const headers = INVENTAIRE_HEADERS;
  const colIndex = headers.indexOf(emplacement) + 1; // 1-based
  if (colIndex <= 0) return { success: false, error: 'Emplacement invalide' };

  const before = Number(item[emplacement]) || 0;
  const after = Math.max(0, before + Number(delta));

  sheet.getRange(item._row, colIndex).setValue(after);
  sheet.getRange(item._row, headers.indexOf('DateMAJ') + 1).setValue(nowIso_());

  logHistorique_({
    CodeBarre: item.CodeBarre,
    Nom: item.Nom,
    Action: action || (delta < 0 ? 'Consommation' : 'Ajout'),
    EmplacementSource: emplacement.replace('Qte',''),
    EmplacementDestination: '',
    QuantiteAvant: before,
    QuantiteApres: after,
    Commentaire: ''
  });

  return { success: true, before: before, after: after };
}

/**
 * Déplace X bouteilles d'un emplacement à un autre
 * emplacementSource / emplacementDest = 'QteCellier' | 'QteLocker' | 'QteFrigo'
 */
function moveBottle(id, emplacementSource, emplacementDest, qte) {
  const sheet = getSheet_(SHEET_INVENTAIRE);
  const items = sheetToObjects_(sheet);
  const item = items.find(it => it.ID === id);
  if (!item) return { success: false, error: 'Bouteille introuvable' };

  const headers = INVENTAIRE_HEADERS;
  const srcCol = headers.indexOf(emplacementSource) + 1;
  const dstCol = headers.indexOf(emplacementDest) + 1;
  if (srcCol <= 0 || dstCol <= 0) return { success: false, error: 'Emplacement invalide' };

  const srcBefore = Number(item[emplacementSource]) || 0;
  const dstBefore = Number(item[emplacementDest]) || 0;
  qte = Number(qte);

  if (qte > srcBefore) return { success: false, error: 'Quantité insuffisante à la source' };

  const srcAfter = srcBefore - qte;
  const dstAfter = dstBefore + qte;

  sheet.getRange(item._row, srcCol).setValue(srcAfter);
  sheet.getRange(item._row, dstCol).setValue(dstAfter);
  sheet.getRange(item._row, headers.indexOf('DateMAJ') + 1).setValue(nowIso_());

  logHistorique_({
    CodeBarre: item.CodeBarre,
    Nom: item.Nom,
    Action: 'Deplacement',
    EmplacementSource: emplacementSource.replace('Qte',''),
    EmplacementDestination: emplacementDest.replace('Qte',''),
    QuantiteAvant: srcBefore,
    QuantiteApres: srcAfter,
    Commentaire: qte + ' bouteille(s) déplacée(s) vers ' + emplacementDest.replace('Qte','')
  });

  return { success: true };
}

// ============================================================
// HISTORIQUE / JOURNAL
// ============================================================
function logHistorique_(entry) {
  const sheet = getSheet_(SHEET_HISTORIQUE);
  sheet.appendRow([
    generateId_(),
    nowIso_(),
    entry.CodeBarre || '',
    entry.Nom || '',
    entry.Action || '',
    entry.EmplacementSource || '',
    entry.EmplacementDestination || '',
    entry.QuantiteAvant != null ? entry.QuantiteAvant : '',
    entry.QuantiteApres != null ? entry.QuantiteApres : '',
    entry.Commentaire || ''
  ]);
}

function getHistorique(limit) {
  const sheet = getSheet_(SHEET_HISTORIQUE);
  const items = sheetToObjects_(sheet);
  items.sort((a, b) => new Date(b.DateHeure) - new Date(a.DateHeure));
  return limit ? items.slice(0, limit) : items;
}

// ============================================================
// RÉCUPÉRATION DE DONNÉES SAQ (scraping)
// ============================================================

/**
 * Interroge l'API GraphQL interne de SAQ.com (utilisée par leur propre site)
 * pour retrouver la fiche produit correspondant à un code-barres, puis
 * en extrait les métadonnées via le JSON-LD de la page produit.
 *
 * NOTE : cette API n'est pas documentée officiellement par la SAQ — elle a été
 * identifiée par inspection du trafic réseau de leur site. Elle peut cesser de
 * fonctionner si la SAQ change son architecture technique (voir guide de dépannage).
 */
function fetchSAQData(barcode) {
  try {
    const found = getSaqProductFromUpc_(barcode);
    if (!found) {
      // Nouvelle tentative avec une variante du code (avec/sans zéro initial)
      let altBarcode = null;
      if (barcode.length === 13) altBarcode = '0' + barcode;
      else if (barcode.length === 14 && barcode.charAt(0) === '0') altBarcode = barcode.substring(1);

      const altFound = altBarcode ? getSaqProductFromUpc_(altBarcode) : null;
      if (!altFound) {
        return { success: false, error: 'Aucun produit trouvé pour ce code-barres sur SAQ.com' };
      }
      return fetchProductDetails_(altFound.productUrl, barcode);
    }
    return fetchProductDetails_(found.productUrl, barcode);

  } catch (err) {
    Logger.log('Exception fetchSAQData : ' + err.toString());
    return { success: false, error: 'Erreur lors de la requête SAQ : ' + err.toString() };
  }
}

/**
 * Appelle l'API GraphQL Adobe Commerce du site SAQ pour retrouver
 * l'URL de la fiche produit correspondant à un code-barres (UPC/EAN).
 * Retourne null si aucun produit n'est trouvé.
 */
function getSaqProductFromUpc_(upc) {
  const endpoint = 'https://catalog-service.adobe.io/graphql';

  const query = `
    query productSearch(
      $phrase: String!
      $pageSize: Int
      $currentPage: Int = 1
      $filter: [SearchClauseInput!]
      $sort: [ProductSearchSortInput!]
      $context: QueryContextInput
    ) {
      productSearch(
        phrase: $phrase
        page_size: $pageSize
        current_page: $currentPage
        filter: $filter
        sort: $sort
        context: $context
      ) {
        total_count
        items {
          product { sku name canonical_url }
          productView { sku name inStock url urlKey }
        }
      }
    }
  `;

  const body = {
    query,
    variables: {
      phrase: String(upc),
      pageSize: 6,
      currentPage: 1,
      filter: [{ attribute: 'visibility', in: ['Search', 'Catalog, Search'] }],
      sort: [],
      context: {
        customerGroup: '356a192b7913b04c54574d18c28d46e6395428ab',
        userViewHistory: []
      }
    }
  };

  const response = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      'Accept': '*/*',
      'Accept-Language': 'fr-CA,fr;q=0.9,en;q=0.8',
      'Origin': 'https://www.saq.com',
      'Referer': 'https://www.saq.com/',
      'x-api-key': '7a7d7422bd784f2481a047e03a73feaf',
      'Magento-Customer-Group': '356a192b7913b04c54574d18c28d46e6395428ab',
      'Magento-Environment-Id': '2ce24571-9db9-4786-84a9-5f129257ccbb',
      'Magento-Store-Code': 'main_website_store',
      'Magento-Store-View-Code': 'fr',
      'Magento-Website-Code': 'base'
    },
    payload: JSON.stringify(body)
  });

  const text = response.getContentText();
  Logger.log('Code réponse GraphQL : ' + response.getResponseCode());

  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    Logger.log('Réponse GraphQL non-JSON : ' + text.substring(0, 500));
    return null;
  }

  const item = json && json.data && json.data.productSearch && json.data.productSearch.items
    ? json.data.productSearch.items[0] : null;
  if (!item) {
    Logger.log('Aucun item GraphQL pour UPC ' + upc);
    return null;
  }

  const saqCode = (item.productView && (item.productView.urlKey || item.productView.sku))
    || (item.product && item.product.sku);
  const productUrl = (item.productView && item.productView.url)
    || ('https://www.saq.com/fr/' + saqCode);

  Logger.log('Produit trouvé via GraphQL : ' + productUrl);
  return { upc: String(upc), saqCode: String(saqCode), productUrl: productUrl };
}

/**
 * Charge la page produit SAQ à l'URL donnée et en extrait les métadonnées
 * via le JSON-LD (voir parseSaqProductHtml_ / extractLabelValue_ plus bas).
 */
function fetchProductDetails_(productUrl, originalBarcode) {
  const productResponse = UrlFetchApp.fetch(productUrl, {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1'
    }
  });

  const prodCode = productResponse.getResponseCode();
  Logger.log('Code réponse fiche produit : ' + prodCode + ' (' + productUrl + ')');
  if (prodCode !== 200) {
    return { success: false, error: 'SAQ a répondu avec le code ' + prodCode + ' (fiche produit : ' + productUrl + ')' };
  }

  const productHtml = productResponse.getContentText();
  const data = parseSaqProductHtml_(productHtml, productUrl);
  Logger.log('Données extraites : ' + JSON.stringify(data));

  if (!data.Nom) {
    return { success: false, error: 'Produit trouvé (' + productUrl + ') mais impossible d\'en extraire les données' };
  }
  data.CodeBarre = originalBarcode;
  return { success: true, data: data };
}

/**
 * Fonction de test à exécuter manuellement dans l'éditeur Apps Script
 * (sélectionner "testSAQ" dans le menu déroulant puis Exécuter),
 * pour diagnostiquer précisément le scraping via Affichage → Journaux d'exécution.
 */
function testSAQ() {
  const result = fetchSAQData('08410310602757');
  Logger.log('RÉSULTAT FINAL : ' + JSON.stringify(result, null, 2));
}

/**
 * Parse le HTML d'une fiche produit SAQ pour en extraire les métadonnées.
 * Utilise des regex simples + JSON-LD si disponible (plus fiable).
 */
function parseSaqProductHtml_(html, url) {
  const result = {
    Nom: '', Producteur: '', Type: '', Pays: '', Region: '',
    Cepage: '', Millesime: '', Format: '750 ml', PrixSAQ: '',
    ImageURL: '', PastilleGout: ''
  };

  // 1) Essai via JSON-LD (souvent le plus fiable et le plus stable dans le temps)
  const ldJsonMatches = html.match(/<script type="application\/ld\+json">([\s\S]*?)<\/script>/g);
  if (ldJsonMatches) {
    for (const block of ldJsonMatches) {
      try {
        const jsonStr = block.replace(/<script type="application\/ld\+json">/, '').replace(/<\/script>/, '');
        const json = JSON.parse(jsonStr);
        if (json['@type'] === 'Product') {
          result.Nom = json.name || result.Nom;
          result.ImageURL = json.image || result.ImageURL;
          if (json.offers && json.offers.price) {
            result.PrixSAQ = json.offers.price;
          }
          if (json.brand && json.brand.name) {
            result.Producteur = json.brand.name;
          }
        }
      } catch (e) {
        // ignore les blocs JSON-LD invalides
      }
    }
  }

  // 2) Repli / complément via regex sur le titre H1
  if (!result.Nom) {
    const titleMatch = html.match(/<h1[^>]*>([^<]+)<\/h1>/);
    if (titleMatch) result.Nom = titleMatch[1].trim();
  }

  // 3) Prix (si non trouvé via JSON-LD)
  if (!result.PrixSAQ) {
    const priceMatch = html.match(/([0-9]+[.,][0-9]{2})\s*\$/);
    if (priceMatch) result.PrixSAQ = priceMatch[1].replace(',', '.');
  }

  // 4) Champs descriptifs (pays, région, cépage, type, millésime)
  // La SAQ affiche souvent ces infos dans un tableau de caractéristiques
  // avec des libellés du type "Pays", "Région", "Cépage", "Couleur", "Millésime"
  result.Pays = extractLabelValue_(html, ['Pays']);
  result.Region = extractLabelValue_(html, ['Région', 'Region']);
  result.Cepage = extractLabelValue_(html, ['Cépage', 'Cépages']);
  result.Type = extractLabelValue_(html, ['Couleur', 'Type de produit', 'Type']);
  result.Millesime = extractLabelValue_(html, ['Millésime']);
  result.Format = extractLabelValue_(html, ['Format']) || '750 ml';

  return result;
}

/**
 * Cherche un libellé (ex: "Pays") suivi de sa valeur dans le HTML,
 * en gérant plusieurs variantes de structure possibles.
 */
function extractLabelValue_(html, labels) {
  for (const label of labels) {
    // Pattern typique : <span>Label</span> ... <span>Valeur</span>
    const regex = new RegExp(label + '\\s*(?:</[a-z]+>)?\\s*(?:<[^>]+>)*\\s*([A-Za-zÀ-ÿ0-9 ,\'\\-]{2,60})<', 'i');
    const match = html.match(regex);
    if (match && match[1]) {
      return match[1].trim();
    }
  }
  return '';
}

// ============================================================
// STATISTIQUES POUR LE DASHBOARD
// ============================================================
function getStats() {
  const items = getInventaire();
  let totalCellier = 0, totalLocker = 0, totalFrigo = 0;
  const parType = {};

  items.forEach(it => {
    totalCellier += Number(it.QteCellier) || 0;
    totalLocker += Number(it.QteLocker) || 0;
    totalFrigo += Number(it.QteFrigo) || 0;
    const type = it.Type || 'Non classé';
    const qte = (Number(it.QteCellier)||0) + (Number(it.QteLocker)||0) + (Number(it.QteFrigo)||0);
    parType[type] = (parType[type] || 0) + qte;
  });

  return {
    totalBouteilles: totalCellier + totalLocker + totalFrigo,
    totalCellier: totalCellier,
    totalLocker: totalLocker,
    totalFrigo: totalFrigo,
    nombreReferences: items.length,
    parType: parType
  };
}

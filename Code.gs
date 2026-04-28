const FEUILLE_MATRICE = 'matrice';
const FEUILLE_HISTORIQUE = 'Historique Saisies';
const FEUILLE_MODELE_ETIQUETTE = 'Etiquette';
const FEUILLE_SUIVI = 'Suivi'; 
const FEUILLE_ARCHIVE = 'Archive';
const FORMAT_UNIQUE = 'Etiquette'; 

// En-têtes pour la recherche dynamique
const COL_HEADER_ETIQUETTE_NUM = 'N° Étiquette'; 
const COL_HEADER_DATE_HEURE = 'Date/Heure';
const COL_HEADER_LIVRE = 'Livré'; // Supposons que cette case doit être cochée
const COL_HEADER_COMMENTAIRE_NAVETTE = 'Commentaire Navette'; // Nouvelle colonne pour le commentaire

// ======================================================================
// 1. DÉCLENCHEURS ET UI
// ======================================================================
/**
 * Affiche la barre latérale du menu principal (NOUVEAU).
 */
function showMenuPrincipal() {
  Logger.log("Démarrage de showMenuPrincipal."); // Ligne de débogage
    const html = HtmlService.createHtmlOutputFromFile('MenuPrincipal')
        .setWidth(300)
        .setTitle('Menu d\'Actions Rapides');
    SpreadsheetApp.getUi().showSidebar(html);
  Logger.log("Tentative d'ouverture de la barre latérale."); // Ligne de débogage
}
function initialiserApplication() {
    const ui = SpreadsheetApp.getUi();
    
    try {
        // Tente d'accéder aux services nécessitant des autorisations étendues
        const userEmail = Session.getActiveUser().getEmail(); 
        
        // Tente d'accéder aux données de la feuille pour forcer la portée 'Spreadsheet'
        SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange('A1').getValue(); 
        
        ui.alert("✅ Autorisations accordées. L'application est maintenant prête.\n\nL'utilisateur détecté est : " + userEmail);
        
        // Relancer le menu principal après l'initialisation
        showMenuPrincipal(); 

    } catch (e) {
        // Si cette erreur se produit, c'est que l'autorisation n'a pas été accordée
        ui.alert("⚠️ Autorisations requises. Veuillez cliquer sur 'Vérifier les autorisations' et suivre les étapes de Google pour accorder l'accès.");
    }
}
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏷️ Actions')
      .addItem('Ouvrir Menu Principal', 'showMenuPrincipal')
      .addItem('Nouvelle Saisie', 'showDialog') 
      .addItem('Réimpression par N°', 'ouvrirDialogueReimpression')
      .addItem('Modifier Saisie', 'ouvrirDialogueModification')
      .addItem('Contrôler la Dépose', 'ouvrirDialogueControleDepose') // NOUVEAU MENU
      .addItem('🔒Archiver & Nettoyer (Calendrier)', 'openDatePickerDialog')
      .addToUi();
}

function showDialog() {
  const template = HtmlService.createTemplateFromFile('Formulaire');
  template.lieuxDepart = getLieuxDepart();
  template.lieuxLitiges = getLitiges();
  template.TypeOption = getType();
  template.ManutentionOptions = getManutention();
  
  const html = template.evaluate().setWidth(400).setHeight(600).setTitle('Saisie des Informations');
  SpreadsheetApp.getUi().showSidebar(html);
}

function ouvrirDialogueReimpression() {
  const template = HtmlService.createTemplateFromFile('FormulaireReimpression');
  template.etiquettesEnCours = getEtiquettesEnCours(); 
  
  const html = template.evaluate().setWidth(400).setHeight(600).setTitle('Réimpression');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Ouvre la barre latérale pour modifier une étiquette.
 * @param {string|null} [etiquetteNumToLoad=null] - Numéro d'étiquette à charger automatiquement.
 * @param {string|null} [litigeValueToSet=null] - Valeur à définir pour le champ Litige ('' pour 'Conforme').
 */
function ouvrirDialogueModification(etiquetteNumToLoad = null, litigeValueToSet = null) {
  const template = HtmlService.createTemplateFromFile('FormulaireModification');
  template.lieuxDepart = getLieuxDepart();
  template.lieuxLitiges = getLitiges();
  template.TypeOption = getType();
  template.ManutentionOptions = getManutention();
  template.etiquettesEnCours = getEtiquettesEnCours();

  template.etiquetteToLoad = etiquetteNumToLoad; 
  template.litigeValueToSet = litigeValueToSet;

  const html = template.evaluate().setWidth(400).setHeight(600).setTitle('Modification de Saisie');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Ouvre la barre latérale pour le contrôle de la dépose.
 */
function ouvrirDialogueControleDepose() {
    const template = HtmlService.createTemplateFromFile('ControleDepose');
    template.lieuxDeposeConnus = getLieuxDepart(); 
    template.etiquettesEnCours = getEtiquettesEnCours();

    const html = template.evaluate().setWidth(400).setHeight(450).setTitle('Contrôle de la Dépose');
    SpreadsheetApp.getUi().showSidebar(html);
}


function onEdit(e) {
  // Cette fonction est laissée vide ou absente, car nous utilisons le menu.
}


// ======================================================================
// 2. LOGIQUE D'ARCHIVAGE ET CONTRÔLE
// ======================================================================

/**
 * Traite le contrôle de dépose, valide le lieu, coche la case "Livré" 
 * et met à jour le commentaire dans la feuille Suivi.
 */
function controlerDeposeEtLivrer(formData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetHistorique = ss.getSheetByName(FEUILLE_HISTORIQUE);
    const sheetSuivi = ss.getSheetByName(FEUILLE_SUIVI);
    
    // Constantes d'index dans la ligne d'historique (0-based)
    const INDEX_ETIQUETTE_NUM_HIST = 0; // N° Étiquette (Colonne A)
    const INDEX_DEPOSE_HISTORIQUE = 8;  // Lieu de dépose (Colonne I)
    
    // Constantes de Colonnes (1-based) pour la mise à jour
    const COL_DATE_VALIDATION = 14;     // Colonne N
    const COL_USER_VALIDATION = 15;     // Colonne O

    const etiquetteNum = formData.etiquetteNum.trim();
    const gareDepose = formData.gareDepose.trim();
    const commentaireNavette = formData.commentaireNavette || '';
    
    // Récupérer les index de colonnes dans Suivi via les en-têtes
    const colEtiquetteNum = getHeaderColumnIndex(sheetSuivi, COL_HEADER_ETIQUETTE_NUM);
    const colLivre = getHeaderColumnIndex(sheetSuivi, COL_HEADER_LIVRE);
    const colCommentaireNavette = getHeaderColumnIndex(sheetSuivi, COL_HEADER_COMMENTAIRE_NAVETTE);

    if (!sheetHistorique || !sheetSuivi) {
        return { success: false, message: "Feuille Historique ou Suivi introuvable." };
    }
    
    // 1. Récupérer les données de l'Historique (Lieu de dépose attendu)
    // NOTE: getLastRow() peut être lente. Si la portée est toujours 13 colonnes, c'est mieux.
    const dataHistorique = sheetHistorique.getDataRange().getValues(); 

    const rowIndexHistorique = dataHistorique.findIndex((row, index) => 
        index > 0 && String(row[INDEX_ETIQUETTE_NUM_HIST]).trim() === etiquetteNum
    );

    if (rowIndexHistorique === -1) {
        return { success: false, message: `Erreur: N° Etiquette "${etiquetteNum}" introuvable dans l'historique.` };
    }
    
    const rowHistorique = dataHistorique[rowIndexHistorique];
    const rowNumberHistorique = rowIndexHistorique + 1; // Index de ligne réel (1-based)
    const lieuDeposeAttendu = String(rowHistorique[INDEX_DEPOSE_HISTORIQUE]).trim(); 
    
    // a) Trouver la ligne réelle dans Suivi
    if ([colEtiquetteNum, colLivre, colCommentaireNavette].includes(-1)) {
        return { success: false, message: `Erreur: Colonnes "N° Étiquette", "Livré" ou "Commentaire Navette" introuvable(s) dans la feuille Suivi. Vérifiez les en-têtes.` };
    }
    
    // On réutilise la même logique de recherche que la vôtre
    const dataSuivi = sheetSuivi.getRange(2, colEtiquetteNum, sheetSuivi.getLastRow() - 1, 1).getValues();
    const rowIndexSuivi = dataSuivi.findIndex(row => String(row[0]).trim() === etiquetteNum);

    if (rowIndexSuivi === -1) {
        return { success: false, message: `Erreur: N° Etiquette "${etiquetteNum}" introuvable dans la feuille Suivi.` };
    }

    const rowNumberSuivi = rowIndexSuivi + 2; 

    // 2. Vérification de la correspondance des lieux (MAJUSCULE et TRIM pour la robustesse)
    if (gareDepose.trim().toUpperCase() !== lieuDeposeAttendu.trim().toUpperCase()) {
      
      // La vérification a échoué. On enregistre quand même le commentaire dans Suivi
      if (commentaireNavette) {
        const rangeCommentaire = sheetSuivi.getRange(rowNumberSuivi, colCommentaireNavette);
        const oldCommentaire = String(rangeCommentaire.getValue() || '').trim();
        const newCommentaire = oldCommentaire ? `${oldCommentaire} | ${commentaireNavette}` : commentaireNavette;
        rangeCommentaire.setValue(newCommentaire);
      }
      return { success: false, message: `ERREUR DE DÉPOSE: La gare saisie ("${gareDepose}") ne correspond pas au lieu attendu ("${lieuDeposeAttendu}").` };
    }

    // --- CONTRÔLE OK : PROCÉDER À LA LIVRAISON ---

    // 3. Mise à jour de la feuille Suivi (inchangée)
    
    // a) Écriture de la coche Livré (cellule unique)
    sheetSuivi.getRange(rowNumberSuivi, colLivre).setValue(true);
    
    // b) Gestion du Commentaire Navette (lecture, concaténation, écriture)
    if (commentaireNavette) {
        const rangeCommentaire = sheetSuivi.getRange(rowNumberSuivi, colCommentaireNavette);
        const oldCommentaire = String(rangeCommentaire.getValue() || '').trim();
        const newCommentaire = oldCommentaire ? `${oldCommentaire} | ${commentaireNavette}` : commentaireNavette;
        rangeCommentaire.setValue(newCommentaire);
    }

    // 4. Mise à jour de la feuille Historique Saisies (COLONNES N et O UNIQUEMENT)
    
    const user = Session.getActiveUser().getEmail(); 
    const dateValidation = new Date();
    
    // Écriture uniquement sur les colonnes N (14) et O (15) (2 colonnes)
    sheetHistorique.getRange(rowNumberHistorique, COL_DATE_VALIDATION, 1, 2).setValues([
        [
            dateValidation,        // Col N (Date de validation)
            user                   // Col O (Utilisateur de validation)
        ]
    ]);
    
    // NOTE: Si vous voulez enregistrer le Commentaire Navette dans la colonne L de l'Historique, 
    // vous devez faire une écriture séparée sur la colonne L (ou combiner les setValues).
    // Pour l'instant, seul N et O sont mis à jour comme demandé.

    return { success: true, message: `✅ Dépose confirmée à ${gareDepose}. L'étiquette est marquée comme Livrée.` };
}

function archiverAnciennesSaisies() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const sheetHistorique = ss.getSheetByName(FEUILLE_HISTORIQUE);
    const sheetSuivi = ss.getSheetByName(FEUILLE_SUIVI);
    let sheetArchive = ss.getSheetByName(FEUILLE_ARCHIVE);

    if (!sheetHistorique) {
        ui.alert(`Erreur: La feuille "${FEUILLE_HISTORIQUE}" est introuvable.`);
        return;
    }

    const dateSeuil = new Date();
    dateSeuil.setMonth(dateSeuil.getMonth() - 1);
    
    if (!sheetArchive) {
        sheetArchive = ss.insertSheet(FEUILLE_ARCHIVE);
        const historiqueHeaders = sheetHistorique.getRange(1, 1, 1, 13).getValues()[0];
        sheetArchive.appendRow(historiqueHeaders);
    }

    const lastRowHistorique = sheetHistorique.getLastRow();
    const dataHistorique = sheetHistorique.getRange(2, 1, lastRowHistorique - 1, 13).getValues();
    
    let archivesCount = 0;
    
    for (let i = dataHistorique.length - 1; i >= 0; i--) {
        const rowHistorique = dataHistorique[i];
        const rowNumberHistorique = i + 2; 

        const dateCreation = rowHistorique[1]; 
        const etiquetteNum = String(rowHistorique[0]).trim();
        
        const isOld = dateCreation instanceof Date && dateCreation < dateSeuil;
        
        if (isOld) {
            sheetArchive.appendRow(rowHistorique);
            sheetHistorique.deleteRow(rowNumberHistorique);
            
            if (sheetSuivi) {
                const colEtiquetteSuivi = getHeaderColumnIndex(sheetSuivi, COL_HEADER_ETIQUETTE_NUM);
                if (colEtiquetteSuivi !== -1) {
                    const dataSuiviRange = sheetSuivi.getRange(2, colEtiquetteSuivi, sheetSuivi.getLastRow() - 1, 1);
                    const valuesSuivi = dataSuiviRange.getValues().map(row => String(row[0]).trim());
                    const indexSuivi = valuesSuivi.findIndex(num => num === etiquetteNum);
                    
                    if (indexSuivi !== -1) {
                        const rowNumberSuivi = indexSuivi + 2;
                        sheetSuivi.deleteRow(rowNumberSuivi);
                    }
                }
            }
            
            archivesCount++;
        }
    }
    
    ui.alert(`Archivage terminé. ${archivesCount} entrées archivées et supprimées des feuilles de travail.`);
}


// ======================================================================
// 4. LOGIQUE D'IMPRESSION/MODIFICATION ET UTILITAIRES (INCHANGÉES)
// ======================================================================

function leverLitigeEtOuvrirModification() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSuivi = ss.getSheetByName(FEUILLE_SUIVI);
    
    if (!sheetSuivi) {
        ui.alert(`Erreur: La feuille de suivi "${FEUILLE_SUIVI}" est introuvable.`);
        return;
    }

    const result = ui.prompt(
        'Lever un Litige et Modifier',
        'Veuillez saisir le numéro d\'étiquette (ex: ETQ#000001) :',
        ui.ButtonSet.OK_CANCEL
    );

    if (result.getSelectedButton() !== ui.Button.OK) {
        return;
    }

    const etiquetteNum = result.getResponseText().trim();
    if (!etiquetteNum || !etiquetteNum.startsWith('ETQ#')) {
        ui.alert('Erreur : Numéro d\'étiquette invalide.');
        return;
    }

    const colLitige = getHeaderColumnIndex(sheetSuivi, 'Lever le litige');
    const colEtiquetteNum = getHeaderColumnIndex(sheetSuivi, 'N° Étiquette');
    
    if (colLitige !== -1 && colEtiquetteNum !== -1) {
        const dataRange = sheetSuivi.getRange(2, colEtiquetteNum, sheetSuivi.getLastRow() - 1, 1);
        const values = dataRange.getValues().map(row => row[0]);
        const rowIndex = values.findIndex(num => String(num).trim() === etiquetteNum);

        if (rowIndex !== -1) {
            const rowNumber = rowIndex + 2; 
            sheetSuivi.getRange(rowNumber, colLitige).setValue(true);
        } else {
             ui.alert('Avertissement : Le N° Etiquette n\'a pas été trouvé dans la feuille Suivi, mais le processus continue.');
        }
    } else {
         ui.alert('Avertissement : En-têtes "Lever le litige" ou "N° Étiquette" introuvables. Le processus continue.');
    }
    
    // Ouvrir le formulaire de modification et définir le champ Litige à 'Conforme' (valeur vide '')
    ouvrirDialogueModification(etiquetteNum, ''); 

    ui.alert(`Formulaire de modification prêt pour ${etiquetteNum}. Le champ Litige est réglé sur 'Conforme'.`);
}


function getHeaderColumnIndex(sheet, headerName) {
    if (!sheet) return -1;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const index = headers.findIndex(header => header.trim() === headerName.trim());
    return index !== -1 ? index + 1 : -1;
}

function getOptions(colIndex) {
    const classeur = SpreadsheetApp.getActiveSpreadsheet();
    const feuilleMatrice = classeur.getSheetByName(FEUILLE_MATRICE);
    if (!feuilleMatrice) {
        SpreadsheetApp.getUi().alert(`Erreur : L'onglet "${FEUILLE_MATRICE}" est introuvable. Veuillez le créer.`);
        return [];
    }

    const range = feuilleMatrice.getRange(2, colIndex, feuilleMatrice.getLastRow(), 1); 
    const options = new Set();
    
    range.getValues().forEach(row => {
        const value = String(row[0]).trim();
        if (value.length > 0) {
            options.add(value);
        }
    });

    return Array.from(options).sort();
}

function getLitiges() { return getOptions(1); }
function getLieuxDepart() { return getOptions(2); }
function getManutention() { return getOptions(3); }
function getType() { return getOptions(4); }

function getEtiquettesEnCours() {
    const sheetSuivi = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_SUIVI);
    
    if (!sheetSuivi || sheetSuivi.getLastRow() <= 1) {
        return [];
    }

    const colLivre = getHeaderColumnIndex(sheetSuivi, COL_HEADER_LIVRE); 
    const colEtiquetteNum = getHeaderColumnIndex(sheetSuivi, COL_HEADER_ETIQUETTE_NUM);
    
    if (colEtiquetteNum === -1) {
        return [];
    }
    
    const lastRow = sheetSuivi.getLastRow();
    const dataRange = sheetSuivi.getRange(2, colEtiquetteNum, lastRow - 1, 1);
    const allValues = dataRange.getValues();
    
    const etiquetteList = [];
    
    allValues.forEach((row, index) => {
        const etiquetteNum = String(row[0]).trim();
        
        // Si la colonne Livré existe, vérifier si elle est cochée (TRUE) ou contient une date
        let estLivre = false;
        if (colLivre !== -1) {
            const livreValue = sheetSuivi.getRange(index + 2, colLivre).getValue(); 
            estLivre = livreValue === true || livreValue instanceof Date;
        }
        
        if (etiquetteNum.startsWith('ETQ#') && !estLivre) {
            etiquetteList.push(etiquetteNum);
        }
    });

    return etiquetteList.sort();
}

function getNextEtiquetteNumber() {
  const historiqueSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTORIQUE);
  if (!historiqueSheet || historiqueSheet.getLastRow() <= 1) {
    return "ETQ#000001"; 
  }

  const lastEtiquetteValue = historiqueSheet.getRange(historiqueSheet.getLastRow(), 1).getValue();
  const lastNumberString = String(lastEtiquetteValue).replace("ETQ#", "");
  const nextNumber = parseInt(lastNumberString) + 1;
  
  const paddedNumber = ("000000" + nextNumber).slice(-6); 
  
  return `ETQ#${paddedNumber}`;
}

function genererEtiquette(formData, numeroEtiquette) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName(FEUILLE_MODELE_ETIQUETTE);

    if (!templateSheet) {
        return { success: false, newSheetName: null };
    }
    
    const dateString = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yy HH:mm");
    const newSheetName = `${numeroEtiquette}_${dateString}`;
    
    const newSheet = templateSheet.copyTo(ss).setName(newSheetName);

    // Remplissage des cellules Étiquette
    newSheet.getRange('B4').setValue(numeroEtiquette);
    newSheet.getRange('D2').setValue(`${formData.lieuDepart}`);
    newSheet.getRange('C4').setValue(formData.Motif);
    newSheet.getRange('C5').setValue(`Qt : ${formData.quantite} - ${formData.TypeUnite}`);
    newSheet.getRange('D3').setValue(`${formData.lieuDepose}`);
    newSheet.getRange('C6').setValue(`Commentaire : ${formData.commentaire}`);
    
       // --- CORRECTION : Ajout d'une courte pause avant l'activation ---
 
        ss.setActiveSheet(newSheet); 
        
       

    return { success: true, newSheetName: newSheetName };


}

function activateSheetByName(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetToActivate = ss.getSheetByName(sheetName);
    
    if (sheetToActivate) {
        // Ajout d'une petite pause pour la synchronisation de l'UI
        Utilities.sleep(50);
        ss.setActiveSheet(sheetToActivate);
    }
}

function enregistrerHistorique(formData, numeroEtiquette) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let historiqueSheet = ss.getSheetByName(FEUILLE_HISTORIQUE);
    const utilisateur = Session.getActiveUser().getEmail(); 

    if (!historiqueSheet) {
        historiqueSheet = ss.insertSheet(FEUILLE_HISTORIQUE); 
        historiqueSheet.appendRow([
            "N° Étiquette", "Date/Heure", "Lieu de départ", "Motif", "N° Lot", 
            "N° Série", "Quantité", "TypeUnite", "Lieu de dépose", "Commentaire", 
            "Litige", "Utilisateur", "Version" // 13 colonnes
        ]);
    }

    const rowData = [
        numeroEtiquette, new Date(), formData.lieuDepart, formData.Motif, 
        formData.lot, formData.serie, formData.quantite, formData.TypeUnite, 
        formData.lieuDepose, formData.commentaire, formData.litige,
        utilisateur,
        0 // Version initiale
    ];
    
    historiqueSheet.appendRow(rowData);
}

function imprimer(formData) {
  const numeroEtiquette = getNextEtiquetteNumber();
  
  enregistrerHistorique(formData, numeroEtiquette);
  
  const result = genererEtiquette(formData, numeroEtiquette); 
  const generationMessage = result.success ? "Une feuille d'étiquette temporaire a été générée." : `ATTENTION : Le modèle '${FEUILLE_MODELE_ETIQUETTE}' est introuvable.`;
  
  return {
    message: `Opération terminée. Numéro créé: **${numeroEtiquette}**. ` + generationMessage,
    sheetName: result.newSheetName
    // format n'est plus retourné car non suivi
  };
}

function getEtiquetteDataForModification(numeroEtiquette) {
  const historiqueSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTORIQUE);
  
  if (!historiqueSheet || historiqueSheet.getLastRow() <= 1) {
    return null;
  }
  
  const allValues = historiqueSheet.getRange(2, 1, historiqueSheet.getLastRow() - 1, 13).getValues(); 
  const rowIndex = allValues.findIndex(row => String(row[0]).trim() === numeroEtiquette.trim());

  if (rowIndex === -1) {
    return null;
  }

  const rowData = allValues[rowIndex];
  const rowNumber = rowIndex + 2;
  
  const [
    etiquetteNum, dateHeure, lieuDepart, Motif, lot, serie, quantite, TypeUnite, 
    lieuDepose, commentaire, litige, utilisateur, version 
  ] = rowData;
  
  const dateFormatted = Utilities.formatDate(dateHeure, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");

  const formData = {
    etiquetteNum, dateHeure: dateFormatted, lieuDepart, Motif, lot, serie, quantite, TypeUnite, lieuDepose, 
    commentaire, litige, utilisateur, version, 
    formatImpression: FORMAT_UNIQUE // Ajout du format par défaut pour l'affichage dans le formulaire
  };

  return { formData: formData, row: rowNumber };
}

/**
 * Cœur de la sauvegarde: Incrémente la version et met à jour la ligne.
 */
function saveModification(formData, rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historiqueSheet = ss.getSheetByName(FEUILLE_HISTORIQUE);
  
  if (!historiqueSheet) {
    return { success: false, message: "Feuille d'historique introuvable." };
  }
  
  const newModificationDate = new Date();
  const newModifierUser = Session.getActiveUser().getEmail();

  const oldVersion = parseInt(historiqueSheet.getRange(rowNumber, 13).getValue() || 0); 
  const newVersion = oldVersion + 1;

  const rowData = [
    formData.etiquetteNum,
    newModificationDate, // Nouvelle date/heure
    formData.lieuDepart,
    formData.Motif,
    formData.lot,
    formData.serie,
    formData.quantite,
    formData.TypeUnite,
    formData.lieuDepose,
    formData.commentaire,
    formData.litige,
    newModifierUser, // Nouvel utilisateur
    newVersion // Nouvelle version
  ];
  
  historiqueSheet.getRange(rowNumber, 1, 1, 13).setValues([rowData]);
  
  // Renvoie le numéro d'étiquette et la nouvelle version
  return { success: true, message: `L'étiquette ${formData.etiquetteNum} a été modifiée (v${newVersion}) avec succès.`, etiquetteNum: formData.etiquetteNum, newVersion: newVersion };
}

/**
 * NOUVEAU : Sauvegarde uniquement la modification sans imprimer (Bouton 1)
 */
function saveModificationOnly(formData, rowNumber) {
    const result = saveModification(formData, rowNumber);
    
    if (result.success) {
        return { success: true, message: `✅ Modification enregistrée (v${result.newVersion}). Aucune impression demandée.` };
    }
    return result; // Propagate error
}

/**
 * Gère le workflow de Litige complet (SANS IMPRESSION).
 */
function processerLitigeAndSave(formData, rowNumber) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSuivi = ss.getSheetByName(FEUILLE_SUIVI);
    const etiquetteNum = formData.etiquetteNum;
    
    // 1. COCHER LA CASE DANS LE SUIVI
    if (sheetSuivi) {
        const colLitige = getHeaderColumnIndex(sheetSuivi, 'Lever le litige');
        const colEtiquetteNum = getHeaderColumnIndex(sheetSuivi, 'N° Étiquette');
        
        if (colLitige !== -1 && colEtiquetteNum !== -1) {
            const dataRange = sheetSuivi.getRange(2, colEtiquetteNum, sheetSuivi.getLastRow() - 1, 1);
            const values = dataRange.getValues().map(row => String(row[0]).trim());
            const rowIndex = values.findIndex(num => num === etiquetteNum);
            
            if (rowIndex !== -1) {
                const rowSuivi = rowIndex + 2; 
                sheetSuivi.getRange(rowSuivi, colLitige).setValue(true); // COCHER LA CASE
            }
        }
    }

    // 2. Sauvegarder la modification (Litige est déjà sur "Conforme" dans formData)
    const saveResult = saveModification(formData, rowNumber);

    if (saveResult.success) {
        // IMPORTANT: Ne retourne pas l'objet d'impression, mais un message de succès simple.
        return { success: true, message: `✅ Litige réglé (v${saveResult.newVersion}) et sauvegarde terminée.` };
    }
    return saveResult; // Propagate error
}


function reimprimerParNumero(numeroEtiquette) {
  const historiqueSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEUILLE_HISTORIQUE);
  
  if (!historiqueSheet || historiqueSheet.getLastRow() <= 1) {
    return { success: false, message: "Feuille d'Historique Saisies est introuvable ou vide." };
  }
  
  const allValues = historiqueSheet.getRange(2, 1, historiqueSheet.getLastRow() - 1, 13).getValues(); 
  const rowData = allValues.find(row => String(row[0]).trim() === numeroEtiquette.trim());

  if (!rowData) {
    return { success: false, message: `Numéro d'étiquette "${numeroEtiquette}" introuvable.` };
  }
  
  const [
    etiquetteNum, , lieuDepart, Motif, lot, serie, quantite, TypeUnite, 
    lieuDepose, commentaire, litige, , version
  ] = rowData;
  
  const formData = { lieuDepart, Motif, lot, serie, quantite, TypeUnite, lieuDepose, commentaire, litige };
  
  const result = genererEtiquette(formData, etiquetteNum);

  if (!result.success) {
    return { success: false, message: "Erreur de génération du document temporaire (modèle manquant?)." };
  }

  return { 
    success: true, 
    message: `Étiquette ${etiquetteNum} (v${version}) générée. La feuille temporaire est active.`, 
    sheetName: result.newSheetName
  };
}

/**
 * Supprime la feuille temporaire et active la feuille de Suivi.
 */
function deleteSheetByName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetToDelete = ss.getSheetByName(sheetName);
  const sheetSuivi = ss.getSheetByName(FEUILLE_SUIVI); 

  if (sheetToDelete) {
    ss.deleteSheet(sheetToDelete);
    
    // Activer la feuille de Suivi
    if (sheetSuivi) {
        ss.setActiveSheet(sheetSuivi);
    }
    
    return "La feuille '" + sheetName + "' a été supprimée avec succès. Le panneau va se fermer.";
  }
  return "Erreur : La feuille à supprimer n'a pas été trouvée.";
}

const TEMP_SHEET_PREFIX = 'Etiquette_';
const CLEANUP_DURATION_MS = 60 * 60 * 1000; // 1 heure en millisecondes

/**
 * Nettoie toutes les feuilles d'étiquette temporaires créées il y a plus d'une heure.
 * Cette fonction est destinée à être lancée par un déclencheur basé sur le temps (ex: toutes les heures).
 * NOTE: Dépend du format de nom de feuille Etiquette_YYYYMMDD_HHMMSS.
 */
function deleteOldTemporarySheets() {
  const TEMP_SHEET_PREFIX = 'Etiquette_';
  const CLEANUP_DURATION_MS = 1000*60*60*5; // 5 heure en millisecondes
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentTime = new Date();
  let sheetsDeleted = 0;
    ss.getSheets().forEach(sheet => {
        const sheetName = sheet.getName();
        
        // Cible les feuilles qui commencent par ETQ#
        if (sheetName.startsWith('ETQ#')) {
            const dayStr = sheetName.substring(11, 13);
            const monthStr = sheetName.substring(14, 16);
            const yearShortStr = sheetName.substring(17, 19);
            const hourStr = sheetName.substring(20, 22);
            const minuteStr = sheetName.substring(23, 25);
             // Conversion en nombres
            const day = parseInt(dayStr);
            const month = parseInt(monthStr);
            const year = 2000 + parseInt(yearShortStr);
            const hour = parseInt(hourStr);
            const minute = parseInt(minuteStr);
            const sheetTimestamp = new Date(year, month - 1, day, hour, minute, 0);
          if (currentTime - sheetTimestamp > CLEANUP_DURATION_MS) {
                                ss.deleteSheet(sheet);
                                sheetsDeleted++;

                }
    
         if (sheetsDeleted > 0) {
          Logger.log(`Nettoyage automatique terminé : ${sheetsDeleted} feuilles temporaires ont été supprimées.`);
          }
        }
    })
}

function openSlide() {
  const url = "https://docs.google.com/presentation/d/1eXXca5p3cTITKZ0aKRFMnjnb_3JJe5PhojuzQLrvsWY/edit?slide=id.g25460785f2a_0_875#slide=id.g25460785f2a_0_875";
  
  // Contenu HTML de la boîte de dialogue
  const htmlContent = `
    <p>Veuillez cliquer sur le lien ci-dessous  :</p>
    <p><strong><a href="${url}" target="_blank">${url}</a></strong></p>
    <style>
      body { font-family: sans-serif; }
    </style>
  `;

  // Crée la sortie HTML et définit la taille de la boîte
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(550) 
      .setHeight(180);

  // Affiche la boîte de dialogue modale
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'S2-BD38 - Plans des Lieux de dépose campus AL-aT');
}
function onMyEdit(e) {
  // Définitions des constantes
  const SHEET_NAME_SUIVI = "Suivi";
  const COL_LIVRE = 15; // Colonne O (15ème colonne)
  const COL_Etiquette_SUIVI = 1; // Colonne 1 (supposition que l'étiquette est en col A)
  
  const SHEET_NAME_HISTORIQUE = "Historique Saisies";
  const COL_Etiquette_HISTORIQUE = 1; // Colonne A Etiquette dans Historique Saisies
  const COL_DATE_VALIDATION = 14; // Colonne N (nouvelle colonne)
  const COL_USER_VALIDATION = 15; // Colonne O (nouvelle colonne)
  
  const range = e.range;
  const sheet = range.getSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Vérification des conditions de déclenchement
  if (sheet.getName() !== SHEET_NAME_SUIVI || 
      range.getColumn() !== COL_LIVRE || 
      range.getRow() === 1 || 
      e.value !== 'TRUE') { // On vérifie si la case est cochée (valeur 'TRUE')
    return;
  }
  
  const etq = sheet.getRange(range.getRow(), COL_Etiquette_SUIVI).getValue();
  const historiqueSheet = ss.getSheetByName(SHEET_NAME_HISTORIQUE);

  if (!historiqueSheet) {
    // La feuille d'historique n'existe pas, rien à faire
    return;
  }
  
  const historiqueData = historiqueSheet.getDataRange().getValues();
  
  // 2. Recherche du etq dans l'Historique Saisies
  for (let i = 1; i < historiqueData.length; i++) { // Commence à 1 pour ignorer l'en-tête
    // Compare le etq de la ligne en cours avec le etq de l'événement
    if (historiqueData[i][COL_Etiquette_HISTORIQUE - 1] == etq) {
      
      const targetRow = i + 1; // +1 pour obtenir l'index de ligne réel
      
      // 3. Mise à jour des colonnes de validation
      
      // Récupère l'utilisateur actuel
      const user = Session.getActiveUser().getEmail(); 
      
      // Date et utilisateur de validation
      historiqueSheet.getRange(targetRow, COL_DATE_VALIDATION).setValue(new Date());
      historiqueSheet.getRange(targetRow, COL_USER_VALIDATION).setValue(user);
      
    
      // On arrête la recherche après avoir trouvé et mis à jour
      return;
    }
  }
}
/**
 * Ouvre la boîte de dialogue modale pour la sélection de la date d'archivage.
 */
function openDatePickerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('DateArchivagePicker')
      .setWidth(350)
      .setHeight(250)
      .setTitle('Sélection de la Date Limite');
  SpreadsheetApp.getUi().showModalDialog(html, 'Choisir la Date d\'Archivage');
}

/**
 * Fonction de rappel appelée par le HTML pour démarrer l'archivage avec la date sélectionnée.
 * @param {string} dateString - La date limite sélectionnée (YYYY-MM-DD).
 */
function triggerArchivageWithDate(dateString) {
  // Convertir la chaîne en objet Date
  const dateSeuil = new Date(dateString);
  
  if (isNaN(dateSeuil.getTime())) {
    // Cela ne devrait pas se produire si l'interface utilisateur est bien gérée.
    SpreadsheetApp.getUi().alert('Erreur: La date sélectionnée est invalide.');
    return;
  }
  
  // Appeler la fonction principale d'archivage avec la date valide
  archiverEtNettoyerHistorique(dateSeuil);
}


/**
 * Archive les lignes de l'Historique Saisies si :
 * 1. La case 'Livré' (colonne N) est cochée dans la feuille Suivi.
 * 2. La date de création de l'étiquette dans l'Historique (colonne B) est antérieure à la date spécifiée par l'utilisateur.
 * * @param {Date} dateSeuil - La date limite d'archivage fournie par l'utilisateur via le calendrier.
 */
function archiverEtNettoyerHistorique(dateSeuil) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- Constantes ---
    const FEUILLE_HISTORIQUE = 'Historique Saisies';
    const FEUILLE_ARCHIVE = 'Archive';
    const FEUILLE_SUIVI = 'Suivi';
    
    // Colonnes dans l'Historique Saisies (1-based)
    const COL_ETIQUETTE_HIST = 1;         // Colonne A
    const COL_DATE_HEURE_HIST = 2;        // Colonne B (Date de création de l'étiquette)
    const COL_FIN_HISTORIQUE = 15;        // Colonne O (Utilisateur de validation)

    // Colonnes dans le Suivi (1-based)
    const COL_ETIQUETTE_SUIVI = 1;        // Colonne A (N° Étiquette)
    const COL_LIVRE_SUIVI = 15;           // Colonne O (Case cochée/Valeur TRUE)
    const COL_FIN_DECALAGE_SUIVI = 16;    // Colonne P (Fin de la plage de recherche dans Suivi)

    // --- Vérification et Initialisation des Feuilles ---
    const sheetHistorique = ss.getSheetByName(FEUILLE_HISTORIQUE);
    const sheetSuivi = ss.getSheetByName(FEUILLE_SUIVI);
    let sheetArchive = ss.getSheetByName(FEUILLE_ARCHIVE);

   

    if (!sheetArchive) {
        sheetArchive = ss.insertSheet(FEUILLE_ARCHIVE);
        const historiqueHeaders = sheetHistorique.getRange(1, 1, 1, COL_FIN_HISTORIQUE).getValues()[0];
        sheetArchive.appendRow(historiqueHeaders);
    }
    
    // Ajuster la date au début du jour pour que la comparaison soit précise ('avant le 00:00:00 de ce jour')
    dateSeuil.setHours(0, 0, 0, 0);
    
    // Récupération des données nécessaires
    const lastRowSuivi = sheetSuivi.getLastRow();
    // Nous lisons de A à O pour avoir l'étiquette (A) et la case cochée (N)
    const dataSuivi = sheetSuivi.getRange(2, 1, lastRowSuivi - 1, COL_FIN_DECALAGE_SUIVI).getValues();
    const lastRowHistorique = sheetHistorique.getLastRow();
    // getDataRange() est coûteux. Si la feuille a des milliers de lignes vides, c'est mieux de faire getRange(1, 1, lastRowHistorique, COL_FIN_HISTORIQUE)
    const dataHistorique = sheetHistorique.getDataRange().getValues(); 
    
    let archivesCount = 0;

    // TRAVAIL DU BAS VERS LE HAUT (i = Index de ligne réel 1-based)
    for (let i = lastRowHistorique; i >= 2; i--) { 
        const rowData = dataHistorique[i - 1]; // Conversion en index 0-based pour le tableau
        
        const etiquetteNum = String(rowData[COL_ETIQUETTE_HIST - 1]).trim();
        const dateCreation = rowData[COL_DATE_HEURE_HIST - 1];
        
        // --- VÉRIFICATION DE LA LIVRAISON (Dans le Suivi) ---
        
        const rowIndexSuiviArray = dataSuivi.findIndex(row => String(row[COL_ETIQUETTE_SUIVI - 1]).trim() === etiquetteNum);
        
        let isDelivered = false;
        let rowNumberSuivi = -1;

        if (rowIndexSuiviArray !== -1) {
            // Col N (15) a l'index 13 dans le tableau dataSuivi
            const livreValue = dataSuivi[rowIndexSuiviArray][COL_LIVRE_SUIVI - 1]; 
            
            // La case est cochée si la valeur est le booléen TRUE
            isDelivered = (livreValue === true); 
            
            rowNumberSuivi = rowIndexSuiviArray + 2; // Ligne réelle
        } else {
            // Si l'étiquette n'est plus dans le Suivi, elle n'est pas archivée.
            continue; 
        }

        // --- VÉRIFICATION DE L'ANCIENNETÉ (Dans l'Historique) ---
        // Vrai si la date de création est une date ET est strictement antérieure au seuil défini par l'utilisateur
        const isOld = dateCreation instanceof Date && dateCreation < dateSeuil;
        
        // --- VÉRIFICATION FINALE et EXÉCUTION ---
        if (isDelivered && isOld) {
            
            // SUPPRESSION DE LA LIGNE DANS LA FEUILLE SUIVI (si trouvée)
            if (rowNumberSuivi !== -1) {
                // NOTE : Cette suppression doit être faite avant la suppression dans Historique,
                // car elle dépend de l'index du tableau dataSuivi qui n'est pas affecté ici.
                sheetSuivi.deleteRow(rowNumberSuivi);
            }
            
            // ARCHIVAGE ET SUPPRESSION DE L'HISTORIQUE
            
            // Copie vers la feuille Archive (A à O)
            const rowRangeHistorique = sheetHistorique.getRange(i, 1, 1, COL_FIN_HISTORIQUE);
            const valuesToArchive = rowRangeHistorique.getValues()[0];
            sheetArchive.appendRow(valuesToArchive);
            
            // Suppression de la ligne de l'Historique. On part de i=max vers 2.
            sheetHistorique.deleteRow(i);
            
            archivesCount++;
        }
    }
    
    // Afficher le résultat
   
}

function getWorkingDaysAgo(n) {

  let date = new Date();

  let workingDaysCount = 0;



  // Initialise la date du jour à 00:00:00

  date.setHours(0, 0, 0, 0); 

  

  while (workingDaysCount < n) {

    // Recule d'un jour calendaire

    date.setDate(date.getDate() - 1); 

    

    // Récupère le jour de la semaine (0=Dimanche, 6=Samedi)

    const dayOfWeek = date.getDay(); 

    

    // Un jour ouvré est entre Lundi (1) et Vendredi (5)

    if (dayOfWeek >= 1 && dayOfWeek <= 5) {

      workingDaysCount++;

    }

  }

  return date;

}



/**

 * Fonction principale à déclencher. Elle calcule le seuil J-2 jours ouvrés 

 * et lance l'archivage avec cette date.

 */

function lancerArchivageQuotidien() {

    // Calcul de la date seuil : 2 jours ouvrés en arrière

    const J_MOINS_2_OUVRES = 2;

    const dateSeuilCalculee = getWorkingDaysAgo(J_MOINS_2_OUVRES);

    

    // Appel de la fonction d'archivage existante

    archiverEtNettoyerHistorique(dateSeuilCalculee);

    

    // NOTE sur les jours fériés : L'exclusion des jours fériés nécessite 

    // d'intégrer un service de calendrier externe (Google Calendar ou API) 

    // ce qui complexifie le script et nécessite des autorisations. 

    // Cette fonction n'exclut que les week-ends.

}

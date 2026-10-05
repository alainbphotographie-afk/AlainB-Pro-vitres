/* AlainB Pro - API Google Sheets
   1) Remplacez CHANGEZ-MOI par un mot de passe de votre choix.
   2) Déployer > Nouveau déploiement > Application Web
      Exécuter en tant que : Moi / Accès : Tout le monde
   3) Copiez l'URL /exec dans l'application (bouton Cloud). */
const CLE_SECRETE = 'CHANGEZ-MOI';
const ONGLETS = ['Clients','Chantiers','Devis','Interventions','Factures','Relances','Parametres'];

function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}

function doGet(e){
  try{
    if(!e.parameter || e.parameter.key!==CLE_SECRETE) return out_({ok:false,error:'Clé invalide'});
    const ss=SpreadsheetApp.getActiveSpreadsheet(), res={};
    ONGLETS.forEach(n=>{
      const sh=ss.getSheetByName(n); res[n]=[];
      if(!sh||sh.getLastRow()<1) return;
      const v=sh.getDataRange().getDisplayValues(), h=v[0];
      for(let i=1;i<v.length;i++){
        if(v[i].every(c=>c==='')) continue;
        const o={}; h.forEach((k,j)=>{if(k) o[k]=v[i][j];}); res[n].push(o);
      }
    });
    return out_({ok:true,sheets:res,date:new Date().toISOString()});
  }catch(err){return out_({ok:false,error:String(err)});}
}

function doPost(e){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const p=JSON.parse(e.postData.contents);
    if(p.key!==CLE_SECRETE) return out_({ok:false,error:'Clé invalide'});
    if(p.action!=='push') return out_({ok:false,error:'Action inconnue'});
    const s=p.sheets||{};
    if(!s.Clients||s.Clients.length<2) return out_({ok:false,error:'Refusé : onglet Clients vide (protection contre l\'écrasement)'});
    const ss=SpreadsheetApp.getActiveSpreadsheet(), counts={};
    ONGLETS.forEach(n=>{
      const aoa=s[n]; if(!aoa||!aoa.length) return;
      let sh=ss.getSheetByName(n)||ss.insertSheet(n);
      sh.clearContents();
      const w=Math.max.apply(null,aoa.map(r=>r.length));
      const m=aoa.map(r=>{while(r.length<w)r.push('');return r;});
      const rg=sh.getRange(1,1,m.length,w);
      rg.setNumberFormat('@').setValues(m);
      sh.setFrozenRows(1); sh.getRange(1,1,1,w).setFontWeight('bold');
      counts[n]=m.length-1;
    });
    SpreadsheetApp.flush();
    return out_({ok:true,counts:counts,date:new Date().toISOString()});
  }catch(err){return out_({ok:false,error:String(err)});}
  finally{lock.releaseLock();}
}

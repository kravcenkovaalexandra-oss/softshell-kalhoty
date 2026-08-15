const SHEET_NAME = "Objednávky";
const RECIPIENT_EMAIL = "kravcenkovaalexandra@gmail.com";

function setup(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(SHEET_NAME);
  if(!sh) sh=ss.insertSheet(SHEET_NAME);
  if(sh.getLastRow()===0){
    sh.appendRow([
      "Datum","Kluk/Holka","Velikost","Základní cena","Varianta",
      "Barva základ","Barva nápletu","Vzor kolen/kapes",
      "Reflexní prvky","Příplatek reflex","Zimní/Jarní",
      "Doručení","Doprava","Ks","Cena/ks","CELKEM",
      "Jméno","Telefon","E-mail","Poznámka"
    ]);
    sh.setFrozenRows(1);
  }
}

function doPost(e){
  try{
    setup();
    const d=JSON.parse(e.postData.contents);
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    sh.appendRow([
      new Date(),d.sex,d.size,d.basePrice,d.variant,d.baseColor,d.cuffColor,
      d.pattern,d.reflective,d.reflectivePrice,d.season,d.delivery,d.shipping,
      d.quantity,d.unitPrice,d.total,d.name,d.phone,d.email,d.note||""
    ]);

    const body =
      "Nová objednávka softshellových kalhot\n\n"+
      "Kluk/Holka: "+d.sex+"\n"+
      "Velikost: "+d.size+"\n"+
      "Varianta: "+d.variant+"\n"+
      "Barva základ: "+d.baseColor+"\n"+
      "Barva nápletu: "+d.cuffColor+"\n"+
      "Vzor kolen/kapes: "+d.pattern+"\n"+
      "Reflexní prvky: "+d.reflective+"\n"+
      "Zimní/Jarní: "+d.season+"\n"+
      "Doručení: "+d.delivery+"\n"+
      "Počet: "+d.quantity+"\n"+
      "Cena/ks: "+d.unitPrice+" Kč\n"+
      "Doprava: "+d.shipping+" Kč\n"+
      "CELKEM: "+d.total+" Kč\n\n"+
      "Jméno: "+d.name+"\n"+
      "Telefon: "+d.phone+"\n"+
      "E-mail: "+d.email+"\n"+
      "Poznámka: "+(d.note||"");

    MailApp.sendEmail(RECIPIENT_EMAIL,"Nová objednávka softshellových kalhot",body);

    return ContentService.createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
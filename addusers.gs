var fullImpressora = "1wJrOvuHaCLOCe9RXMn9HY-dqKKl-KNxpFaSpEolhN3Q";
var emailConsergeria = "fotocopies@iesperealsius.cat";
var emailCapEstudis = "capestudis@iesperalsius.cat";

function testGeneraEmails()
{ 
  var dict = {"nous":emaiCapEstudis};
  var fullCalcul = "1gCvBaoSleq01xraKYTWI2Jcy_Lm3xNa2hAw1vCfm4rg";
  var i =0;
  for (const [key, value] of Object.entries(dict)) {
      generaEmails(SpreadsheetApp.openById("1gCvBaoSleq01xraKYTWI2Jcy_Lm3xNa2hAw1vCfm4rg").getSheetByName(key),value);
      i = i +1;
      Logger.log("%s - Curs %s Tutor %s generat",i,key, value);
  }
}

function testIdAlumne()
{
  var ui = FormApp.getUi();
  var alumne={};
  alumne.nom="Roger";
  alumne.cognoms="Carol";
  alumne.nivell="ESO";
  alumne.curs="1";
  alumne.email="";
  ui.alert(generateId(alumne));
 alumne.nom="Paula";
  alumne.cognoms="Bartroli";
  alumne.nivell="PROFE";
  alumne.curs="";
  alumne.email="mcolom29@iesperealsius.cat";
  ui.alert(generateId(alumne));
  alumne.nom="Jaume";
  alumne.cognoms="Sanchis";
  alumne.nivell="PROFE";
  alumne.curs="";
  alumne.email="jsanch55@iesperealsius.cat";
  ui.alert(generateId(alumne));
}

function testOuAlumne()
{
  var ui = FormApp.getUi();
  ui.alert(generateOu("ESO","3"));
  ui.alert(generateOu("BATX","2"));
  ui.alert(generateOu("",""));
  ui.alert(generateOu("EE","1"));
  ui.alert(generateOu("CF","2"));
}

function testOmplirColumnes()
{
  var id = "1l_pJ40ugMwUFcr6DPNQ2BWF7zYsWHpu8EgKYGJmm8p4";
  var columnes = omplirColumnes(SpreadsheetApp.openById(id).getSheets()[0]);
  MailApp.sendEmail("coordinacio.tic@iesperealsius.cat", "Prova", JSON.stringify(columnes))
}


function testColumnesOK(){
    fullaCalculESO=SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1exyO4oWEGSbhdDDUDBJn9way_RCK63vZ/edit?usp=sharing&ouid=115695442973019916559&rtpof=true&sd=true").getSheets()[0]
    var columnes = omplirColumnes(fullaCalculESO);    
    var fila = 4;
    var body = "Nom,Cognoms,Nivell,Curs,CorreuExtern,EmailResponsable,Cohort";

    //Generem correus
    while(fila<5)
    {
        var filaAlumne = fullaCalculESO.getRange(fila,1, 1, fullaCalculESO.getLastColumn()).getValues()[0];
        body+= filaAlumne[columnes.Nom] + ",";    
        body+= filaAlumne[columnes.Cognoms] + ",";
        body+= filaAlumne[columnes.Nivell] + ",";
        body+= filaAlumne[columnes.Curs] + ",";
        body+= filaAlumne[columnes.CorreuExtern] + ",";
        body+= filaAlumne[columnes.EmailResponsable] + ",";
        body+= filaAlumne[columnes.Cohort] + ",";
        body+= filaAlumne[columnes.Moodle] + ",";
        fila++;
    }  
  
    MailApp.sendEmail("coordinacio.tic@iesperealsius.cat", "Dades usuaris ", body, {  noReply:true});
    
}

function testNeteja(){

  var text = netejar("àbsçºª");
  text = text + "p";
}
function getCurs()
{
  var data = new Date();
  if (data.getMonth()<6){
    return data.getFullYear()-1;
  }
  else{
    return data.getFullYear();
  }

}
function generateId(alumne)
{
  var id = "";
  var it = 1;
  var cognom = (alumne.cognoms.indexOf(" ")>0 ? alumne.cognoms.split(" ")[0] : alumne.cognoms);
  if (alumne.nivell.localeCompare("PROFE")!=0)
  { id = eliminablancs(alumne.nom.charAt(0).toLowerCase()+cognom.toLowerCase()+getCurs().toString().substring(2));
    var newId = id;
    while(isUser(newId+"@iesperealsius.cat")){
      newId=id+"."+it;
      newId=eliminablancs(newId.toLowerCase());
      it++;
    }  
  } 
  else
  { id = alumne.email.substring(0,alumne.email.indexOf("@"));
    it = 1;
    var newId = id;
    while(isUser(newId+"@iesperealsius.cat")){
      newId=id+it;
      newId=eliminablancs(newId.toLowerCase());
      it++;
    }  
  }  
  return newId;
}

function generateOu(nivell,curs)
{
   if (nivell.localeCompare("ESO")==0)
      return "/Alumnat/"+curs+"ESO";
   if (nivell.localeCompare("BATX")==0)
      return "/Alumnat/"+curs+"BATX";  
   if (nivell.localeCompare("EE")==0)
      return "/Alumnat/EnsenyamentsEsportius";  
    if (nivell.localeCompare("CF")==0)
      return "/Alumnat/CICLES";        
   return "/Professorat";
}

function netejar(text){
  var cercar= new Array('à','á','ä','À','Á','Ä','é','è','ë','É','È','Ë','í','ì','ï','Í','Ì','Ï','ó','ò','ö','Ó','Ò','Ö','ú','ù','ü','Ú','Ù','Ü','Ç','ç','ñ','Ñ','ª','º');
 var subs= new Array('a','a','a','A','A','A','e','e','e','E','E','E','i','i','i','I','I','I','o','o','o','O','O','O','u','u','u','U','U','U','C','c','n','N','a','o');
 text = text.toString();
  for(var i=0;i<cercar.length;i++){
    text=text.replace(new RegExp(cercar[i], 'g'),subs[i]);
  } 
  return text;
}

function eliminablancs(text){
  text=text.replace(new RegExp(' ', 'g'),'');
  return text;
}

function setContrasenya(){
  var length = 8;
  var text = "";
  var possible = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789";

  for( var i=0; i < length; i++)
    text += possible.charAt(Math.floor(Math.random() * possible.length));

  return text;
}

function isUser(id){
  try{
    //Si no el troba llença una excepció
    var user=AdminDirectory.Users.get(id);
    return true;
  }
  catch(e){
    return false;
  }
}

function testGetMail(){
  Logger.log("%s",getMail("Hermegildo","Porres"));
  Logger.log("%s",getMail("Jana","Puerto"));
  Logger.log("%s",getMail("Marc","Juan"));
}

function getMail(nom,cognoms)
{
  try{
    var userList = AdminDirectory.Users.list({
        domain: 'iesperealsius.cat',
        query: "name:'" + nom+" "+cognoms + "'"
    }).users;
  
    if (userList.length==1)
      return userList[0].primaryEmail;  
    return userList.length;
  }
  catch(e){
    return "";
  }
}

function totalEmailsGenerats(fulla,colCorreuInstitut){
  rowTmp=1;
  while(!fulla.getRange(rowTmp,1).isBlank() && !fulla.getRange(rowTmp,colCorreuInstitut).isBlank()){
      rowTmp++;
  } 
  return rowTmp;
}


function crearFullaEmails(usuarisGeneratsId){
    var fullaEmailsGenerats = SpreadsheetApp.openById(usuarisGeneratsId).getActiveSheet();
    fullaEmailsGenerats.getRange(1,1).setValue("Nivell ");
    fullaEmailsGenerats.getRange(1,2).setValue("Curs ");
    fullaEmailsGenerats.getRange(1,3).setValue("Grup ");
    fullaEmailsGenerats.getRange(1,4).setValue("Nom ");
    fullaEmailsGenerats.getRange(1,5).setValue("Cognoms");
    fullaEmailsGenerats.getRange(1,6).setValue("Correu ");
    fullaEmailsGenerats.getRange(1,7).setValue("Password");
    fullaEmailsGenerats.getRange(1,8).setValue("Usuari Moodle");
    fullaEmailsGenerats.getRange(1,8).setValue("Password Moodle");
    return fullaEmailsGenerats;
}

function crearFullaMoodle(id){
    var fulla = SpreadsheetApp.openById(id).getActiveSheet();
    fulla.getRange(1,1).setValue("username");
    fulla.getRange(1,2).setValue("password");
    fulla.getRange(1,3).setValue("firstname");
    fulla.getRange(1,4).setValue("lastname");
    fulla.getRange(1,5).setValue("email");
    fulla.getRange(1,6).setValue("cohort1");
    return fulla;
}

function getColumnaNom(name,valors,ocurrencia){
    var vegada = 1;
    for (var j=0; j < valors.length; j++)
    {
         var nomColumna = valors[j];
         if (nomColumna.localeCompare(name)==0)
         {
             if (ocurrencia == vegada)
             {
               return j;
             }
             else
             {
               vegada++;
             }
         }
    }
    return -1;
}
         
function omplirColumnes(fulla){
    var columnes = {};
    var capcalera =  fulla.getRange(1,1, 1,fulla.getLastColumn()).getValues()[0];
    columnes["Nom"] = getColumnaNom("Nom",capcalera,1);
    columnes["Cognoms"] = getColumnaNom("Cognoms",capcalera,1);
    columnes["CorreuExtern"] = getColumnaNom("Correu extern",capcalera,1);
    columnes["EmailResponsable"] = getColumnaNom("Email responsable",capcalera,1);
    columnes["Nivell"] = getColumnaNom("Nivell",capcalera,1);
    columnes["Curs"] = getColumnaNom("Curs",capcalera,1);
    columnes["Grup"] = getColumnaNom("Grup",capcalera,1);
    columnes["Cohort"] = (getColumnaNom("Cohort",capcalera,1)>-1?getColumnaNom("Cohort",capcalera,1):"");
    columnes["Moodle"] = (getColumnaNom("Moodle",capcalera,1)>-1?getColumnaNom("Moodle",capcalera,1):"");
    columnes["EmailInsti"] = (getColumnaNom("EmailInsti",capcalera,1)>-1?getColumnaNom("EmailInsti",capcalera,1):"");
    columnes["DNI"] = (getColumnaNom("DNI",capcalera,1)>-1?getColumnaNom("DNI",capcalera,1):"");
    return columnes;  
}

function escriureDadesAlumne(fulla,fila,alumne){
  fulla.getRange(fila,1).setValue(alumne.nivell);
  fulla.getRange(fila,2).setValue(alumne.curs);
  fulla.getRange(fila,3).setValue(alumne.grup);
  fulla.getRange(fila,4).setValue(alumne.nom);
  fulla.getRange(fila,5).setValue(alumne.cognoms);
  if (alumne.emailInsti.toString().indexOf("@")>0)
 {
    fulla.getRange(fila,6).setValue(alumne.emailInsti);
    fulla.getRange(fila,7).setValue("Ja el té d'altres cursos");
 }  
  else
  {
    fulla.getRange(fila,6).setValue(alumne.id+"@iesperealsius.cat");
    fulla.getRange(fila,7).setValue(alumne.password);
  }  
  fulla.getRange(fila,8).setValue(alumne.id);
  fulla.getRange(fila,9).setValue(alumne.passwordMoodle);
}

function escriureMoodle(fulla,fila,alumne){
  fulla.getRange(fila,1).setValue(alumne.id);
  fulla.getRange(fila,2).setValue(alumne.passwordMoodle);
  fulla.getRange(fila,3).setValue(alumne.nomRaw);
  fulla.getRange(fila,4).setValue(alumne.cognomsRaw);
  fulla.getRange(fila,5).setValue(alumne.id+"@iesperealsius.cat");
  fulla.getRange(fila,6).setValue(alumne.cohort);
}

function escriureImpressora(fulla,alumne){
  fulla.appendRow([alumne.nom,alumne.cognoms,alumne.id,alumne.dni])
}

function creacioUsuariGoogle(fila,columnes,alumne)
{
    //Creació usuari al GSuite
    //Afegim a l'unitat organitzativa corresponent https://stackoverflow.com/questions/10872122/adding-domain-user-to-organisation-unit-using-google-apps-script
    //Estructura user https://developers.google.com/admin-sdk/directory/v1/reference/users/insert
    if (fila[columnes.CorreuExtern] == null)
    {
      var user = {
        primaryEmail: alumne.id+"@iesperealsius.cat",
        orgUnitPath:alumne.ou, 
        name: {
                givenName:  alumne.nomRaw,
                familyName: alumne.cognomsRaw
        },
        password: alumne.password
      };
    }
    else
    {
       var user = {
        primaryEmail: alumne.id+"@iesperealsius.cat",
        orgUnitPath:alumne.ou, 
        name: {
                givenName: alumne.nomRaw,
                familyName: alumne.cognomsRaw
        },
        password: alumne.password,  
        "emails": [
        {
               "address": fila[columnes.CorreuExtern],
               "type": "home",
              "primary": false
        }
        ]
       };
     } 
     AdminDirectory.Users.insert(user);
     afegirGrup(alumne);
}

function afegirGrup(alumne)
{
  var userEmail = (alumne.emailInsti.indexOf("@")>0?alumne.emailInsti:alumne.id+'@iesperealsius.cat');
  var groupId = alumne.cohort+"@iesperealsius.cat";
  try{

    var group = GroupsApp.getGroupByEmail(groupId);
    var newMember = {email: userEmail, role: "MEMBER"};
    AdminDirectory.Members.insert(newMember, groupId);
    Logger.log(userEmail+" is  in the group");
  }
  catch(e){
    Logger.log(userEmail+" is already in the group");
  }  
}

function testMouOu()
{
  alumne = {};
  alumne.id ="sperez83";
  alumne.ou ="/Professorat"
  moureOu(alumne);
}

function moureOu(alumne)
{
  AdminDirectory.Users.update({
          orgUnitPath: alumne.ou,
        }, alumne.id+"@iesperealsius.cat")
}

function comunicarDades(filaAlumne,columnes,alumne,doc)
{
     // Agafa el model, crea una copia com a temporal, i salva l'Id del document creat
     var docMoodle = "1wquNN6rZJ5B05_WrMv4p1YMt77LgxadcvYTzonN7qfc";
     var docNoMail = "1c0pLP_HPGmDEIkFPGN4Gjd-LU39RhX5pQhbOFWQHlwU";
     var docNoMoodle = "1ZgGBkLbCmXJYxDpbj9Cn6hIlpcfbod_QF97fTRUuV8I";
     var docProfeMoodle = "1PoPqQQ2jl1FH2fLLtQ_sm-pPA-4q68aB7iwovxhCHIM";
     var docProfeNoMoodle = "1cWIL5ti_KsHoNMXwLIurgAmjUjKks5iIGd9-o99NLXA";
     
     var docTemplate = null;
     if (alumne.nivell.localeCompare("PROFE")!=0)
      docTemplate = (alumne.emailInsti.toString().indexOf("@")>0?docNoMail:docMoodle);
     else
      docTemplate = docProfeMoodle;
     var copyId = DriveApp.getFileById(docTemplate).makeCopy('Usuari: '+alumne.id).getId();

     //Responsable 1 
     var copyDoc = DocumentApp.openById(copyId);
     var copyBody = copyDoc.getActiveSection();
     copyBody.replaceText('keyCognoms', alumne.cognomsRaw);
     copyBody.replaceText('keyNom', alumne.nomRaw);
     copyBody.replaceText('keyEmail', alumne.id +"@iesperealsius.cat");
     copyBody.replaceText('keyUser', alumne.id);
     if (alumne.emailInsti.indexOf("@")<0)
       copyBody.replaceText('keyPassword', alumne.password);
     else
       copyBody.replaceText('keyPassword', "Ja el tens d'una estada anterior a l'institut. Si no el recordes envia correu a coordinacio.tic@iesperealsius.cat");
     
     if (alumne.moodle.toString().localeCompare("1")!=0)
       copyBody.replaceText('keyPwdMoodle', alumne.password);
     else
       copyBody.replaceText('keyPwdMoodle', "Ja el tens d'una estada anterior a l'institut. Si no el recordes envia correu a coordinacio.tic@iesperealsius.cat");
     copyBody.replaceText('keyDNI', alumne.dni);
     //Salva i tanca el fitxer temporal
     copyDoc.saveAndClose();
     var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");
     var bodyHtml = "Bones ,<br/> Benvingut/da a l'institut Pere Alsius, a mode de resum els usuaris que necessites en la teva estada són:<br/>";
     bodyHtml +="<table border='1'><tr><td>Servei</td><td>Usuari</td><td>Contrasenya</td></tr>";
     if (alumne.emailInsti.indexOf("@")<0)
         bodyHtml += "<tr><td>Compte Google</td><td>"+alumne.id +"@iesperealsius.cat"+"</td><td>"+alumne.password+"</td></tr>";
     else
         bodyHtml += "<tr><td>Compte Google</td><td colspan='2'> Ja el tens d'una altra estada a l'institut.</td></tr>";
     if (alumne.moodle.toString().localeCompare("1")!=0)
         bodyHtml += "<tr><td>Moodle</td><td>"+alumne.id +"</td><td>"+alumne.password+"</td></tr>";
     else
         bodyHtml += "<tr><td>Moodle</td><td colspan='2'> Ja el tens d'una altra estada a l'institut.</td></tr>";
     bodyHtml += "<tr><td>Impressió</td><td>"+alumne.id +"</td><td>"+alumne.dni+"</td></tr></table>";
     bodyHtml +=" Adjuntat tens més detall pel que fa a les TIC (usuaris i infraestructures) del nostre institut. Qualsevol dubte escriu-nos un correu a coordinacio.tic@iesperealsius.cat .<br/>Salutacions<br/>Coordinació TIC Institut Pere Alsius";
     var body = "Bones , Benvingut/da a l'institut Pere Alsius, adjuntat tens informació sobre les TIC (usuaris i infraestructures) del nostre institut. Qualsevol dubte escriu-nos un correu a coordinacio.tic@iesperealsius.cat . Salutacions Coordinació TIC Institut Pere Alsius";
       
     if (alumne.email.indexOf("@")>0)
     {
      var subject = "Compte correu i moodle de l'institut Pere Alsius " + alumne.nomRaw + " " + alumne.cognomsRaw ;
       MailApp.sendEmail(alumne.email, subject, body, { htmlBody: bodyHtml, noReply:true, attachments: pdf});
     } 
     else
      if (alumne.emailInsti.indexOf("@")>0)
      {
        var subject = "Compte moodle de l'institut Pere Alsius " + alumne.nomRaw + " " + alumne.cognomsRaw ;
        MailApp.sendEmail(alumne.emailInsti, subject, copyDoc.getBody().getText(), { noReply:true});    
      }  
     doc.getBody().appendParagraph(copyDoc.getBody().getText());
     doc.getBody().appendPageBreak();  
     DriveApp.getFileById(copyId).setTrashed(true);
}


function checkColumnes(){
    return true;
}

function generaEmails(fullaCalculESO,supervisor)
{
  var date = new Date();
  var usuarisGeneratsId = SpreadsheetApp.create("Usuaris"+date.getDate()+"_"+(date.getMonth()+1)+"_"+date.getFullYear()+"_"+date.getMinutes()+"_"+date.getHours()).getId(); 
  var fullaEmailsGenerats = crearFullaEmails(usuarisGeneratsId);
  var csvMoodle = SpreadsheetApp.create("Moodle"+date.getDate()+"_"+(date.getMonth()+1)+"_"+date.getFullYear()+"_"+date.getMinutes()+"_"+date.getHours()+"_"+".csv").getId();  
  var fullaMoodle = crearFullaMoodle(csvMoodle);
  var doc = DocumentApp.create('EmailsAlumnat'+date.getDate()+"_"+(date.getMonth()+1)+"_"+date.getFullYear()+"_"+date.getMinutes()+"_"+date.getHours());
  var columnes = omplirColumnes(fullaCalculESO);  
  var textImpressora ="Bones <br/>S'han de crear aquests nous usuaris per la impressora:<br/>";
  if (checkColumnes(columnes))
  {
    var fila = totalEmailsGenerats(fullaEmailsGenerats,5);
    var filaInicial = fila;
    var filaMoodle = fila;
    var filaImpressora = fila;
   
    //Generem correus
    while(!fullaCalculESO.getRange(fila,1).isBlank())
    {
        var filaAlumne = fullaCalculESO.getRange(fila,1, 1, fullaCalculESO.getLastColumn()).getValues()[0];
        var alumne={}
        alumne.nom = netejar(filaAlumne[columnes.Nom]);
        alumne.cognoms = netejar(filaAlumne[columnes.Cognoms]);

        if (filaAlumne[columnes.Cognoms].indexOf(" ")>0)
        {
          var cognom1 = filaAlumne[columnes.Cognoms].substring(0,filaAlumne[columnes.Cognoms].indexOf(" "));
          var cognom2 =  filaAlumne[columnes.Cognoms].substring(filaAlumne[columnes.Cognoms].indexOf(" ")+1);
          var cognoms = cognom1.charAt(0).toUpperCase()+cognom1.slice(1)+" "+cognom2.charAt(0).toUpperCase()    +cognom2.slice(1)
        }  
        else
        {
          var cognoms = filaAlumne[columnes.Cognoms].charAt(0).toUpperCase()+filaAlumne[columnes.Cognoms].slice(1);
        }
        var nom = filaAlumne[columnes.Nom].charAt(0).toUpperCase()+filaAlumne[columnes.Nom].slice(1);
        alumne.nomRaw = nom;
        alumne.cognomsRaw = cognoms;
        alumne.email = netejar(filaAlumne[columnes.CorreuExtern]);
        alumne.nivell = netejar(filaAlumne[columnes.Nivell]);
        alumne.curs = netejar(filaAlumne[columnes.Curs]);
        alumne.moodle = filaAlumne[columnes.Moodle];
        alumne.emailInsti = filaAlumne[columnes.EmailInsti];
        alumne.grup = filaAlumne[columnes.Grup];
        alumne.ou = generateOu(alumne.nivell,alumne.curs);
        alumne.cohort = filaAlumne[columnes.Cohort];
        alumne.dni = filaAlumne[columnes.DNI];
        alumne.dni = alumne.dni.substring(alumne.dni.length-6,alumne.dni.length-1);
        alumne.password = setContrasenya();
        if (alumne.emailInsti.toString().indexOf("@") < 0 )
        {
          alumne.id = eliminablancs(generateId(alumne).toLowerCase());
          creacioUsuariGoogle(filaAlumne,columnes,alumne);        
        } 
        else
        {
          alumne.id = alumne.emailInsti.substring(0,alumne.emailInsti.indexOf("@"));
          afegirGrup(alumne);
          moureOu(alumne);
        }
        if (alumne.moodle.toString().localeCompare("1")!=0)
        {
          alumne.passwordMoodle = alumne.password;
          escriureMoodle(fullaMoodle,filaMoodle,alumne);
          filaMoodle++;
        }  
        else
        {
          alumne.passwordMoodle = "Ja el té d'altres cursos";
        }
        escriureDadesAlumne(fullaEmailsGenerats,fila,alumne);
        if (alumne.nivell.localeCompare("PROFE")==0){
          escriureImpressora(SpreadsheetApp.openById(fullImpressora),alumne);
          textImpressora += alumne.id +" - "+ alumne.dni+" <br/>";
          filaImpressora++;
        }
          
        comunicarDades(filaAlumne,columnes,alumne,doc);
        fila++;
    }  
  
  
   if (fila!=filaInicial)
   {
     SpreadsheetApp.openById(usuarisGeneratsId).addEditor(supervisor);
     SpreadsheetApp.openById(csvMoodle).addEditor(supervisor);
     doc.addEditor(supervisor);
     var text =  "Bones.<br/> En aquest full de càlcul "+SpreadsheetApp.openById(usuarisGeneratsId).getUrl()+" hi ha les dades d'alumnat  de la teva tutoria. Els password i usuaris pel correu (alumnat nouvingut) i Moodle (tots, ja que és nou).<br/><br/> Aquí també tens en format carta per si vols repartir imprès la comunicació d'usuaris i passwords a l'alumnat "+doc.getUrl() +" . <br/> L'alumnat que ja estava a l'institut també ha rebut per email les dades d'accés a Moodle.<br/>Salutacions";
     MailApp.sendEmail(supervisor, "Emails / Moodle alumnat",text.replace("<br/>","\n"), { bcc:"coordinacio.tic@iesperealsius.cat",htmlBody: text, noReply:true});
     MailApp.sendEmail("coordinacio.tic@iesperealsius.cat","Moodle nous usuaris",SpreadsheetApp.openById(csvMoodle).getUrl());
     if (filaImpressora!=filaInicial)
     {
       MailApp.sendEmail(emailConsergeria,"Impressora nous usuaris",textImpressora+"<br/> Recorda tots els usuaris estan a "+SpreadsheetApp.openById(fullImpressora).getUrl()+"<br /> GRÀCIES!",{htmlBody: textImpressora+"<br/> Recorda tots els usuaris estan a "+SpreadsheetApp.openById(fullImpressora).getUrl()+"<br /> GRÀCIES!", noReply:true});
     }  
   }
   else
   {  
     MailApp.sendEmail(Session.getActiveUser().getEmail(), "Alta professorat/alumnat","No hi ha usuaris nous a generar");    
   }
  }
  else
  {
     MailApp.sendEmail(Session.getActiveUser().getEmail(), "Alta professorat/alumnat - Error","Els noms de les columnes no són correctes.");
  }
}

function onFormSubmit(ev){
    var form = FormApp.getActiveForm();
    //Obtenim TOTES les respostes al formulari (el full de càlcul complert)
    var formResponses = form.getResponses();
    //Agafem la darrera resposta al formulari (darrera fila, la que està enviant l'usuari).
    var formResponse = formResponses[formResponses.length-1];
    //Agafem tots els ítems d'aquesta resposta
    var itemResponses = formResponse.getItemResponses();
    var urlFullaCalculESO =  itemResponses[0].getResponse();
    var emailResponsable =  itemResponses[1].getResponse();
    if (urlFullaCalculESO.indexOf("sharing")<0){
       MailApp.sendEmail(Session.getActiveUser().getEmail(), "Alta alumnat - Error","La url entrada no és correcta. Ha de ser la URL que apareix quan comparteixes el full de càlcul on diu copia enllaç. Si us plau fes-ho i torna a enviar el formulari. Gràcies!");
    }
    else
    {
       var spreadSheet = SpreadsheetApp.openByUrl(urlFullaCalculESO);
       if (spreadSheet == null)
       {
         MailApp.sendEmail(Session.getActiveUser().getEmail(), "Alta alumnat - Error","El full de càlcul no està compartit amb l'usuari coordinacio.tic@iesperealsius.cat. Si us plau fes-ho i torna a enviar el formulari. Gràcies!");
       }
       else
       {
         var editors = spreadSheet.getEditors();
         var canEdit = false;
         for(var e=0; e<editors.length; e++)
         {
            if (editors[e].getEmail().localeCompare("admingoogle@insperealsius.cat")==0){
             canEdit = true;
             break;
            }
          } 
          if (!canEdit)
          {
             MailApp.sendEmail(Session.getActiveUser().getEmail(), "Alta alumnat - Error","Has de donar al teu full de càlcul permisos d'edició a l'usuari coordinacio.tic@iesperealsius.cat. Si us plau, fes-ho i torna a enviar el full de càlcul"); 
          }      
          else
          {
            var fullaCalculESO = spreadSheet.getSheets()[0];
            generaEmails(fullaCalculESO,emailResponsable);
          }
       }
    }
}  

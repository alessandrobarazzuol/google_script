/** carica meet By Alessandro Barazzuol
 * @param {range} calendario id Immmetti range del calendario IDDD
 * @param {range} 
 * @customfunction
 */
function carica_meet() {
 
 try
 {
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Olimpiadi");
  var range=sheet.getActiveRange().getA1Notation();

  sheet.getRange(range).setNumberFormat('@');
  var event=sheet.getRange(range).getValues();
 
  var calendarId;

  for(var i=0;i<event.length;i++)
  {

  var nome_evento=event[i][0];
  var ti=event[i][3];
  var tf=event[i][4];
  var di=event[i][1]+"T"+ti+":00";
  var df=event[i][2]+"T"+tf+":00";
  
   var des=event[i][5];
  
var sheetdati=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dati");
  
   const gmt="+01:00";
    calendarId=sheetdati.getRange("A1").getValue();

   var resource = {
    start: {
    
    dateTime: di + gmt,
    
  },
  end: {
   
    dateTime:df+gmt,
   
  },
  attendees: [

  ],
    
    conferenceData: {
        createRequest: {
            requestId: "Evento"+i,
            conferenceSolutionKey: { type: "hangoutsMeet" },
        },
    },
    summary: nome_evento,
    description: des,
};


  
  if(invitati=event[i][6]!="")
   var invitati=event[i][6].split(",");

   
    for(var j=0; j<invitati.length; j++){
     resource.attendees.push({email: invitati[j]});
     }
     Calendar.Events.insert(resource, calendarId,{conferenceDataVersion:1,sendNotifications: false});
   }
 }catch(e)
 {
   Browser.msgBox("Errore: forse non hai selezionato il range");
 }

  }

 
var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dati");
var ur=sheet.getLastRow();
var range=sheet.getRange("K2:L"+ur).getValues();
var assoc={};
for(var j=0;j<range.length;j++)
{
assoc["f"+j.toString()]=[range[j][0],range[j][1]];
}

Object.keys(assoc).forEach(function (name) {
  this[name]= function () {
   sheet.getRange("A1").setValue(assoc[name][1]);
  
 }
});  
function onOpen() {
  getID();

Logger.log(assoc);
Logger.log("\n");
Logger.log(assoc["Educazione civica 1G"]);



  var ui = SpreadsheetApp.getUi();
  var m=ui.createMenu('Carica Meet')
      .addItem('Carica_Meet(Selezionare il range)', 'carica_meet')
      .addSeparator();





Object.keys(assoc).forEach(function(name){
  m.addItem(assoc[name][0],name.toString());
});

      m.addToUi();


}





function getID()
{
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dati");
  
  var elenco=[];
  var i=0;
  var cal=CalendarApp.getAllCalendars();
  cal.forEach(function(value) {
    elenco[i]=[value.getName(),value.getId()];
    
   i++;
 });
    
   i++;
     var range=sheet.getRange("K2:L"+i);
     range.setValues(elenco);

}




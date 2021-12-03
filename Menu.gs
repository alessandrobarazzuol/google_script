/**
 * carica un menu con sottomenu da foglio excel associando ad ogni voce delle funzioni dinamicamente
 */
/*codice senza funzione viene eseguito prima della funzione*/
var assoc={};
var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
  var range=sheet.getDataRange().getValues();
  /*prima creo un array associativo con le voci e sotto-voci e le funzioni associate*/
  var j;

  /*scorro il range e associo l'array associativo assoc*/
   for(var i=1;i<range.length;i++)
   {
      if(range[i][0]!="")
      {
        j=i;
        var tmp=[];
        while(j<range.length && range[j][1] !="")
        {
          tmp.push(range[j][1]);
          j++;
          
        }
         assoc[range[i][0]]=tmp;
          i=j;
          
      }

  }
    /*scorre l'array associativo assoc */
    Object.keys(assoc).forEach(function(name){

    /*associa le funzioni alle voci del menu principali*/  
    this[name]=function(){
    Browser.msgBox(name);
    }
      /*scorre ogni sotto menu e associa le funzioni alle voci del submenu*/
      assoc[name].forEach(function(subm) {
      this[subm]=function(){
      Browser.msgBox(subm);
      } 
      });
  });


  /*carica il menu*/     
  function Carica_Menu() {


  
  //Logger.log(assoc);

  /*adesso creo il menu*/
   var ui = SpreadsheetApp.getUi();
  
   var menu=ui.createMenu('Menu');
   var submenu;
   Object.keys(assoc).forEach(function(name){
  
    if(assoc[name]=="")
      {
      menu.addItem(name,name)
      menu.addSeparator();
      }
 
    else
      {
      submenu=ui.createMenu(name);
      Object.keys(assoc[name]).forEach(function(id){
      submenu.addItem(assoc[name][id],assoc[name][id]);
      });
      menu.addSubMenu(submenu);
    
      }
 
   });
    
   menu.addToUi();



}

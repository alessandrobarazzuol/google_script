/**
 * Un array associativo è un array fatto in questo modo
 *    var assc={
 *      nome_1="valore",
 *      nome_2="valore",
 *      ecc
 *      }
 * Costruisce un array associativo da un range: la prima colonna diventa il nome e il resto i valori
 * se il range è formato da due colonne la prima è il nome la seconda il valore, altrimenti
 * se il range è formato da piu colonne, la prima diventa il nome e le restanti sono i valori sotto
 * forma di array
 */


function RangetoAssociativeArray() {
  var sheet= SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var range=sheet.getRange("A2:C3").getValues();
  /*array associativo*/
  var assc={};



  for(var i=0;i<range.length;i++)
  {
    var arg=[];
    for(var j=1;j<range[i].length;j++)
    arg.push(range[i][j]);

    assc[range[i][0]]=arg;

  } 

  Logger.log(assc);

  /*visualizzazione array*/
  Object.keys(assc).forEach(name=>{
  Logger.log("Nome campo "+name)
  Logger.log("Contenuto campo "+assc[name][0])
  Logger.log("Descrizione "+assc[name][1])
  }
  );


} 













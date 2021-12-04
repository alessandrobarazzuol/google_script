






/*cerca una datain un array di date*/
function cerca(array,cosa)
{

  for(var i=0;i<array.length;i++)
  {
    
    if(Utilities.formatDate(new Date(array[i]),"GMT+1","dd/MM/yyyy").valueOf()==cosa.valueOf())
    {
          return true;
    }

  }
  return false;

}
/*cerca il giorno della settimana*/
function cercag(array,cosa)
{
for(var i=0;i<array.length;i++)
  {
    
    if(array[i]==cosa && array[i]!="")
    {
          return true;
    }

  }
  return false

}
/**
 * Riempi: riempie un range di date consecutive da di a df scartando i giorni indicati (0,1,2,...6) e le 
 * festivita indicate
 * @param {String} din  data iniziale sotto forma di stringa yyyy/MM/dd
 * @param {String} dfi data finale sotto forma di stringa yyyy/MM/dd
 * @param {Range} fest range delle festivita es H1:H10 NO STRIGHE
 * @param {Range} festvi range dove si trovano i giorni festivi o liberi da togliere es F1:F8 NO STRINGHE 
 * ---> mettere nel range numeri  0 PER domenica,1=lunedi, .....6=sabato. Esempio per togliere Domenica e 
 * Sabato mettere 0 e 6
 * @param {[String]} nome_foglio   nome del foglio opzionale se omesso viene messo il foglio attivo
 * @customfunction
 *  
 */

function Riempi(din,dfi,fest,festvi,nome_foglio)
{
  var foglio;
  if ( arguments.length === 4 ) 
  {
  foglio= SpreadsheetApp.getActiveSheet();
  }
  else
  {
 foglio= SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome_foglio);
  }

var di=new Date(din); 
var df=new Date(dfi);

var utc1 = Date.UTC(di.getFullYear(), di.getMonth(), di.getDate());
var utc2 = Date.UTC(df.getFullYear(), df.getMonth(), df.getDate());
var conta_giorni=Math.floor((utc2 - utc1)/(1000 * 60 * 60 * 24));

var dati=new Array();
dati=[];
for( var i=0;i<=conta_giorni;i++)
{
var tmp=utc1+1000*60*60*24*i;
tmpp=Utilities.formatDate(new Date(tmp),"GMT+1","dd/MM/yyyy");
if(cerca(fest,tmpp)==false && cercag(festvi,new Date(tmp).getDay())==false){
dati.push(tmpp);
}

}
  return dati;
}





function mySleep (sec)
{
 
  Utilities.sleep(sec*1000);
 
}


function onOpen() {
  

  var body=DocumentApp.getActiveDocument().getBody();
  var text=body.editAsText();
  
  for (var i = 0; i<10; i++) {
        text.appendText(i.toString()+" ");
       DocumentApp.getActiveDocument().saveAndClose();
        body = DocumentApp.getActiveDocument().getBody();
        text = body.editAsText();
        mySleep(0,2);
    }
  
}

/**
 * startScript(),keepRunning(),stopScript()
 */
function startScript() {
    if (keepRunning()) return "OK";
    do {
        soluzioni();
    } while (keepRunning());
    return "OK";
}

function keepRunning() {
    var status = PropertiesService.getScriptProperties().getProperty("run") || "OK";
    return status === "OK" ? true : false;
}

function stopScript() {
    PropertiesService.getScriptProperties().setProperty("run", "STOP");
    return "Kill Signal Issued";
}
/**
 *
 *
 * onOpen function : crea il menu
 */
function onOpen() {
    PropertiesService.getScriptProperties().setProperty("run", "STOP");

    var ui = DocumentApp.getUi();

    ui.createMenu("Soluzioni").addItem("Visualizza_Soluzione_Verifica", "startScript").addItem("Reset", "stopScript").addSeparator().addToUi();
}

function go() {
    PropertiesService.getScriptProperties().setProperty("run", "STOP");
    PropertiesService.getScriptProperties().setProperty("go", "OK");
}

/**
 * Lettura delle soluzioni
 */
function soluzioni() {
    PropertiesService.getScriptProperties().setProperty("run", "OK");
    var docURL = DocumentApp.openByUrl("https://docs.google.com/open?id=1vs76XOvgGY9hRU1FXsA4zVuFZrfol6GkRgNJkSCHpes");
    var bodys = docURL.getBody().getText();
    var tables = docURL.getBody().getTables();
    var j = 0;
    var k = 0;
    var body = DocumentApp.getActiveDocument().getBody();
    body.clear();
    var text = body.editAsText();
    var as = false;
    text.setForegroundColor("#ffffff");
    var y = 0;
    var style = {};
    style[DocumentApp.Attribute.FONT_SIZE] = 92;
    text.appendText("").setAttributes(style);
    while (y < 50) {
        if (!keepRunning()) return 0;

        text = body.editAsText();
        text.editAsText().setText("Boot..." + parseInt((y * 100) / 49) + "%");
        DocumentApp.getActiveDocument().saveAndClose();
        body = DocumentApp.getActiveDocument().getBody();

        text = body.editAsText();
        y += 1;
    }

    var style = {};
    style[DocumentApp.Attribute.FONT_SIZE] = 11;
    text.replaceText("Boot...100%", "Start By ALESSANDRO BARAZZUOL").setAttributes(style);
    text.appendText("\n").setBackgroundColor("#000000");

    for (var i = 0; i < bodys.length; i++) {
        if (!keepRunning()) return 0;

        text.appendText(bodys[i].toString());

        DocumentApp.getActiveDocument().saveAndClose();

        body = DocumentApp.getActiveDocument().getBody();

        text = body.editAsText();

        try {
            var position = DocumentApp.getActiveDocument().newPosition(body.getChild(body.getNumChildren() - 1), 1);
            DocumentApp.getActiveDocument().setCursor(position);
        } catch (e) {
            //gestione
        }
    }

    var a = 20;
    var b = 0;
    text.appendText("\nI passaggi del ciclo sono  i seguenti :\n");
    text = body.editAsText();
    while (a != b && a > b) {
        if (!keepRunning()) return 0;
        text.appendText("a=a-b (" + a + " - " + b + ") = " + (a - b) + "\n");

        a = a - b;

        b = b + 1;
        text.appendText("(b)" + (b - 1) + "++ = (" + b + ")\n");
        text.appendText(a + ">" + b + "\n");
        if (a < b) text.appendText(a + ">" + b + " NOOOOOOO\n");
        else text.appendText(a + ">" + b + "..SI\n");

        text.appendText("\n-------------------\n");
        DocumentApp.getActiveDocument().saveAndClose();

        body = DocumentApp.getActiveDocument().getBody();

        mySleep(6);

        try {
            var position = DocumentApp.getActiveDocument().newPosition(body.getParent(), 1);
            DocumentApp.getActiveDocument().setCursor(position);
        } catch (e) {}

        text = body.editAsText();
    }
    text.appendText("Alla fine a vale " + a + "\n");

    return 0;
}

function mySleep(sec) {
    Utilities.sleep(sec * 1000);
}

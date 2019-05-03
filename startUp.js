#targetengine 'conditions';

var myEventHandler2 = function(ev) {
    
    condSwitcher();
}

var myEventHandler1 = function(ev) {

    app.addEventListener('afterSelectionAttributeChanged', myEventHandler2);
}

app.addEventListener('afterOpen', myEventHandler1);

function condSwitcher() {

    var doc = app.activeDocument;
    var edition = doc.activeEdition;
    var florida = app.activeDocument.conditions.item('Florida');
    var sni = app.activeDocument.conditions.item('SNI');
    var national = app.activeDocument.conditions.item('National');

    switch (edition) {

        case 'National':

            florida.visible = false;
            sni.visible = false;
            national.visible = true;
            break;

        case 'SNI':

            florida.visible = false;
            sni.visible = true;
            national.visible = false;
            break;

        case 'Florida':

            florida.visible = true;
            sni.visible = false;
            national.visible = false;
            break;

        default:

            florida.visible = false;
            sni.visible = true;
            national.visible = false
            break;
    }
}
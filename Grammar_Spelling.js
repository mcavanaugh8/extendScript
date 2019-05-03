/* beautify preserve:start */
#include '/Volumes/BEC/15 User Files/INDD SCRIPTS/primitive/main.js';
/* beautify preserve:end */


var doc = app.activeDocument;

fontFix(doc);
clearOverrides(doc);




function fontFix(doc) {

    if (doc.name.match(/g1/i)) {

        var grade = '1';

    } else if (doc.name.match(/g2/i)) {

        var grade = '2';

    } else if (doc.name.match(/g3/i)) {

        var grade = '3';

    } else if (doc.name.match(/g4/i)) {

        var grade = '4';

    } else if (doc.name.match(/g5/i)) {

        var grade = '5';
    }

    var paraStyles = doc.paragraphStyles;
    var p = paraStyles.length;

    while (p--) {

        if (paraStyles[p].appliedFont.name.match(/Verdana/)) {

            // alert('1')

            paraStyles[p].appliedFont = 'Verdana';
        }
    }

    // var charStyles = doc.characterStyleGroups[1].characterStyles.everyItem().getElements();
    var charStyles = doc.characterStyles;
    var c = charStyles.length;

    while (c--) {

        //     if (c !== 0) {
        try {
            if (charStyles[c].appliedFont.match(/Verdana/)) {

                // alert('2')

                charStyles[c].appliedFont = 'Verdana';
            }
        } catch (e) {}
    }
    // }

    try {

        var actTxt = doc.paragraphStyles.item('Activity-txt GR-' + grade);
        actTxt.appliedFont = 'BECAG (T1)';
        actTxt.fontStyle = 'Heavy';
        actTxt.pointSize = 16;
        actTxt.leading = 41;
    } catch (e) {}

    try {

        var actTxt = doc.paragraphStyles.item('Activity-txt GR-' + grade + '-1st-Line');
        actTxt.appliedFont = 'BECAG (T1)';
        actTxt.fontStyle = 'Heavy';
        actTxt.pointSize = 16;
        actTxt.leading = 41;

    } catch (e) {}

    try {
        var descr = doc.paragraphStyles.item('BLM_Description_txt_GR' + grade);
        descr.appliedFont = 'ITC Avant Garde Gothic Std';
        descr.fontStyle = 'Medium';
        descr.pointSize = 16;
        descr.leading = 25;
    } catch (e) {}

    try {
        var a = doc.paragraphStyles.item('BLM-WordBox-txt_GR' + grade);
        a.appliedFont = 'ITC Avant Garde Gothic Std';
        a.fontStyle = 'Medium';
        a.pointSize = 16;
        a.leading = 25;
    } catch (e) {}


    try {
        var nameDate = doc.paragraphStyles.item('Name_Date');
        nameDate.appliedFont = 'Verdana';
        nameDate.fontStyle = 'Regular';
    } catch (e) {}

    try {
        replacer(doc, '\\d\\d*\\.', '$0', 'Activity-txt GR-' + grade, false, false, 'Activity_txt-Bold', false);
    } catch (e) {}


}

function clearOverrides(doc) {

    var doc = app.activeDocument;

    var allstories = doc.stories;
    var a = allstories.length;

    while (a--) {

        var paras = allstories[a].paragraphs;
        var p = paras.length;

        while (p--) {


            paras[p].clearOverrides();


        }
    }

}
/* beautify ignore:start */
#include '/Volumes/BEC/15 User Files/INDD SCRIPTS/primitive/main.js';
/* beautify ignore:end */


var documents = app.documents.everyItem().getElements();
var d = documents.length;

while (d--) {
    notifications(false);
    main();
    notifications(true);

}

function main() {

    // var documents = app.documents.everyItem().getElements();
    // var d = documents.length;
    // notifications(false);

    // while (d--) {
    var doc = app.activeDocument;
    var articles = doc.managedArticles;
    var a = articles.length;

    // var files = getfiles();

    while (a--) {

        articles[a].checkOut();
        // untagger(doc);
        doc.conditions.everyItem().remove();
        var b = doc.conditions.add();
        b.name = 'Florida';
        var c = doc.conditions.add();
        c.name = 'National';
        var d = doc.conditions.add();
        d.name = 'SNI';
        run(doc);
        articles[a].checkIn();

    }

    doc.save();
    doc.close();
}



function run(files) {

    var doc = app.activeDocument;

    if (!doc.characterStyles.item('nat_standards').isValid) {

        var natStyle = doc.characterStyles.item('standards').duplicate();
        natStyle.name = 'nat_standards';

    }

    // replacer(doc, '[A-Z][A-Z]*\\.\\d\\.\\d\\d*[a-z]*', '$0', false, false, false, 'standards', false);
    // replacer(doc, '(K|\\d)\\.[A-Z][A-Z]*\\.\\d.\\d\\d*[a-z]*', '$0', false, false, false, 'standards', false);



    try {
        natter(doc);
        if (doc.characterStyles.item('Standards_Grey').isValid) {

            replacer(doc, '(K|\\d)\\.[A-Z][A-Z]*\\.\\d.\\d[a-z]*', '$0', 'learning_targets_Bullets', false, false, 'standards', false);
            replacer(doc, '.*', '$0', false, 'Standards_Grey', false, 'standards', false);

        } else {
            try {
                replacer(doc, '(K|\\d)\\.[A-Z][A-Z]*\\.\\d.\\da-z]*', '$0', 'learning_targets_Bullets', false, false, 'standards', false);
            } catch (e) {

                alert('1');
                replacer(doc, '(K|\\d)\\.[A-Z][A-Z]*\\.\\d.\\d[a-z]*', '$0', 'learning_targets_Bullets', false, false, 'nat_standards', false);


            }


        }

    } catch (e) {alert('2');};


    try {
        natter(doc);
        if (doc.characterStyles.item('Standards_Grey').isValid) {

            replacer(doc, '(K|\\d)\\.[A-Z][A-Z]*\\.\\d.\\d[a-z]*', '$0', 'student_objectives_Bullets', false, false, 'standards', false);
            replacer(doc, '.*', '$0', false, 'Standards_Grey', false, 'standards', false);

        } else {
            try {
                replacer(doc, '(K|\\d)\\.[A-Z][A-Z]*\\.\\d.\\d[a-z]*', '$0', 'student_objectives_Bullets', false, false, 'standards', false);
            } catch (e) {
                alert('3');
                replacer(doc, '(K|\\d)\\.[A-Z][A-Z]*\\.\\d.\\d[a-z]*', '$0', 'student_objectives_Bullets', false, false, 'nat_standards', false);

            }

        }

    } catch (e) {alert('4');};



    try {
        // applyConditions('.*', false, 'student_objectives_Bullets', doc.conditions[0].name);
        // applyConditions('.*', false, 'student_objectives_A-hd', doc.conditions[0].name);
        applyConditions('.*', 'standards', false, 'Florida');
        applyConditions('.*', 'nat_standards', false, 'National');
    } catch (e) { alert(e + '\r' + e.line) }
}




function findStandards(target, standardMatch, withCstyle) {

    app.findGrepPreferences.findWhat = standardMatch;
    app.findGrepPreferences.appliedCharacterStyle = withCstyle;

    var found = target.findGrep();
    var f = found.length;

    while (f--) {

        var nat = getStandards(found[f].contents);
        // alert(found[f].contents);
        // alert(nat[1]);
        // found[f].insertionPoints[-1].applyCharacterStyle = app.activeDocument.characterStyles.item('nat_stands');
        found[f].insertionPoints[-1].contents = ' ' + nat[1];

    }

    app.findGrepPreferences = NothingEnum.nothing;

}


function natter(target) {

    app.findGrepPreferences.findWhat = '[A-Z][A-Z]*\\.\\d\\.\\d[a-z]*';
    app.findGrepPreferences.appliedCharacterStyle = 'standards';

    var found = target.findGrep();
    var f = found.length;

    while (f--) {

        // alert('NAT Standard: \n' + found[f].contents);
        found[f].appliedCharacterStyle = 'nat_standards';
    }x2x2

    app.findGrepPreferences = NothingEnum.nothing;

}

function spaceDeletor(target) {

    app.findGrepPreferences.findWhat = '\\s(RL|L|W|SL)';
    app.changeGrepPreferences.changeTo = '$1';
    target.changeGrep();

    app.findGrepPreferences = NothingEnum.nothing;


}

function commaAdder(target) {

    app.findGrepPreferences.findWhat = '\\.\\d*[A-Z][A-Z]*';
    app.findGrepPreferences.appliedCharacterStyle = 'nat_standards';
    app.changeGrepPreferences.changeTo = '$1';
    target.changeGrep();

    app.findGrepPreferences = NothingEnum.nothing;

}


function applyConditions(targetText, withCstyle, withPstyle, newCondition) {

    var doc = app.activeDocument;

    if (targetText) {
        app.findGrepPreferences.findWhat = targetText;
    }
    if (withCstyle) {

        app.findGrepPreferences.appliedCharacterStyle = withCstyle;
    }

    if (withPstyle) {
        app.findGrepPreferences.appliedParagraphStyle = withPstyle;
    }

    var found = doc.findGrep();
    var f = found.length;

    while (f--) {

        // alert(found[f]);

        found[f].applyConditions(doc.conditions.item(newCondition));
    }

    app.findGrepPreferences = NothingEnum.nothing;

}

// function applyConditions2(targetText, withCstyle, withPstyle, newCondition) {

//     app.findGrepPreferences.findWhat = targetText;

//     if (withCstyle) {

//         app.findGrepPreferences.appliedCharacterStyle = withCstyle;
//     }

//     if (withPstyle) {
//         app.findGrepPreferences.appliedParagraphStyle = withPstyle;
//     }

//     var found = doc.findGrep();
//     var f = found.length;

//     while (f--) {

//         // alert(found[f]);

//         found[f].applyConditions(newCondition);
//     }

//     app.findGrepPreferences = NothingEnum.nothing;

// }


function getFiles() {

    notifications(false);

    var configfile = File(app.activeScript.parent.fsName + '/config.txt');
    configfile.open('r');
    var str = configfile.read();
    var files = str.split('\n');

    for (var f = 0, length = files.length; f < length; f++) {

        safe_open(files[f]);

    }

    notifications(true);
}

function safe_open(filename) {

    if (app.documents.count() > 5) {

        app.documents.everyItem().close(SaveOptions.no);
    }

    var query = ['Name, ' + filename, 'Type, Layout'];
    var results = app.queryObjects(query);

    var id = !!results.match(/<\d+/) ? results.match(/<\d+/)[0].replace(/</g, '') : false;

    if (!!id) {

        app.openObject(id, true);

    } else {

        alert('id returned false, file not found for filename: ' + filename);
        throw 'exiting program';
    }
}
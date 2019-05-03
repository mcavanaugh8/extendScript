/* beautify preserve:start */
#include '/Volumes/BEC/15 User Files/INDD SCRIPTS/primitive/main.js';
/* beautify preserve:end */

var source = Folder.selectDialog();
var files = source.getFiles('*.xml');
var f = files.length;


while (f--) {

    if (files[f].name.match(/w1/i)) {
        // alert('1');
        app.open(File(app.activeScript.parent.fsName + '/templateW1.indt'));
    } else if (files[f].name.match(/w6/i)) {
        // alert('2');
        app.open(File(app.activeScript.parent.fsName + '/templateW6.indt'));

    } else {
        // alert('3');
        app.open(File(app.activeScript.parent.fsName + '/template.indt'));
    }
    var doc = app.activeDocument;
    // doc.importXML(File('/Users/mcavanaugh/Desktop/StyleMaps for W2D/Writers_Workshop/tabled/WW_template_W2_5-tabled.xml'));
    doc.importXML(File(files[f].fsName));
    var xml = doc.xmlElements[0];
    pacifier(xml);
    namer();
    // metaStyles(files[f]);
    meta(xml);
    elLessonPlacer(xml);
    // pdPlacer(xml);
    if (!files[f].name.match(/w1|w6/i)) {
        topic_placer(xml, ['teach', 'lesson']);
        topic_placerWW(xml, []);
        topic_placerTeach(xml, []);
        topic_placerLesson(xml, []);
    } else {
        topic_placer(xml, ['teach', 'lesson']);
        topic_placerLesson(xml, []);
        topic_placerW1(xml, []);
    }
    eldTabler(doc);
    lessonTabler(doc);
    lookFors(doc);
    // elIconSmall(xml);
    woody_artspecker(xml, doc);
    doc.save(File(files[f].fsName.replace('.xml', '.indd')));
    cleanup(doc);
    untagger();
    doc.save();
    // doc.close();
}

function woody_artspecker(xml, doc) {

    app.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
    doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

    var lessons = xml.evaluateXPathExpression('//learningBECContent');
    var l = lessons.length;

    while (l--) {

        pacifier(lessons[l]);
        var artspecs = lessons[l].evaluateXPathExpression('//data[@name="artspec"]');
        artspecs.sort();
        var page = lessons[l].evaluateXPathExpression('//page');
        page = page[0].xmlAttributes.item('value').value.replace(/\s/g, '');

        var frame = doc.pages.item(page).textFrames.add();
        var x1 = doc.pages.item(page).bounds[0] + 0.025;
        var y1;
        var x2 = doc.pages.item(page).bounds[0] + 2;
        var y2;

        if (page % 2 == 0) {

            y1 = doc.pages.item(page).bounds[1] - 0.5;
            y2 = doc.pages.item(page).bounds[1] - 3;

        } else {

            y1 = doc.pages.item(page).bounds[1] + 3;
            y2 = doc.pages.item(page).bounds[1] + 0.5;
        }


        frame.geometricBounds = [x1, y1, x2, y2];
        frame.name = 'artspec';
        frame.label = 'artspec';
        frame.fillColor = '285 process';

        for (var c = 0; c < artspecs.length; c++) {

            frame.insertionPoints[-1].contents = artspecs[c].xmlAttributes.item('value').value + '\r\r';
        }
    }

    app.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
    doc.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
}

function addRegex(pstyle, cstyle, expression) {

    var target = app.activeDocument.paragraphStyles.item(pstyle);
    var applyThis = app.activeDocument.characterStyles.item(cstyle);
    var newgrep = target.nestedGrepStyles.add()
    newgrep.grepExpression = expression;
    newgrep.appliedCharacterStyle = applyThis;
    return true;
}


function elIconSmall(xml) {

    var elIconSmall = doc.pageItems.item('el_icon_small');
    var placeHere = xml.evaluateXPathExpression('//data[@name="EL_icon"]/preceding-sibling::node()[1]');

    app.select(NothingEnum.NOTHING);
    app.select(elIconSmall);
    app.copy();

    for (var t = 0; t < placeHere.length; t++) {

        app.select(placeHere[t].xmlContent.insertionPoints[-2]);
        app.paste();
    }

    app.select(NothingEnum.NOTHING);

    var paras = doc.paragraphStyles;
    var p = paras.length;

    for (var x = 1; x < p; x++) {

        var target = doc.paragraphStyles[x];

        var newgrep = target.nestedGrepStyles.add()
        newgrep.grepExpression = '~a\\s*\\r';
        newgrep.appliedCharacterStyle = doc.characterStyles.item('el_icon');

    }
}

function lookFors(doc) {

    var lookForsIcon = doc.pageItems.item('look_fors_icon');
    var lookFors = doc.pageItems.item('look_fors');



    var allstories = doc.stories;
    var s = allstories.length;

    while (s--) {

        if (allstories[s].contents.match(/look fors/i)) {

            // alert('1');

            var paras = allstories[s].paragraphs;
            var p = paras.length;

            while (p--) {

                if (paras[p].contents.match(/look fors/i)) {

                    // alert('2');

                    app.findGrepPreferences.findWhat = 'LOOK';
                    app.findGrepPreferences.appliedParagraphStyle = 'look_fors_A-hd';
                    app.changeGrepPreferences.changeTo = 'LK';
                    lookFors.changeGrep();

                    app.select(lookForsIcon);
                    app.copy();
                    app.select(paras[p].insertionPoints[2]);
                    app.paste();
                }
            }
        }
    }
}

function lessonTabler(doc) {

    var frames = doc.textFrames;
    var f = frames.length;

    while (f--) {

        if (!frames[f].name.match(/eld/i)) {

            if (frames[f].parentStory.tables[0].isValid) {

                var docTables = frames[f].parentStory.tables;
                var t = docTables.length;

                for (var x = 0; x < t; x++) {

                    var cells = docTables[x].cells;
                    var c = cells.length;

                    while (c--) {

                        cells[c].topEdgeStrokeWeight = .5;
                        cells[c].topEdgeStrokeTint = 100;
                        cells[c].leftEdgeStrokeWeight = .5;
                        cells[c].leftEdgeStrokeTint = 100;
                        cells[c].rightEdgeStrokeWeight = .5;
                        cells[c].rightEdgeStrokeTint = 100;
                        cells[c].bottomEdgeStrokeWeight = .5;
                        cells[c].bottomEdgeStrokeTint = 100;
                        cells[c].fillColor = 'Paper';
                        // cells[c].strokeColor = doc.swatches.item('Black');


                        cells[c].width = 0.7451 + 'in';
                        cells[c].height = 0.0417 + 'in';


                        docTables[x].cells[0].fillColor = 'week1';
                        docTables[x].cells[0].paragraphs[0].appliedParagraphStyle = 'ELD_table_head';

                        docTables[x].cells[1].fillColor = 'week1';
                        docTables[x].cells[1].paragraphs[0].appliedParagraphStyle = 'ELD_table_head';

                        docTables[x].cells[2].fillColor = 'week1';
                        docTables[x].cells[2].paragraphs[0].appliedParagraphStyle = 'ELD_table_head';

                        app.findGrepPreferences.findWhat = '\\r(?=$)';
                        app.changeGrepPreferences.changeTo = '';
                        cells[c].changeGrep()
                    }
                }
            }
        }
    }
}

function eldTabler(doc) {


    if (doc.paragraphStyles.item('el_lesson_table_Body-txt').isValid) {

        var tableBody = doc.paragraphStyles.item('el_lesson_table_Body-txt');
        tableBody.appliedFont = 'Jotting'
        tableBody.fontStyle = 'Bold';
        tableBody.pointSize = 10;
        tableBody.leading = 16;
    }

    if (!doc.paragraphStyles.item('ELD_table_head').isValid) {

        var tableHead = doc.paragraphStyles.add();
        tableHead.name = 'ELD_table_head';
        tableHead.appliedFont = 'Jotting'
        tableHead.fontStyle = 'Bold';
        tableHead.pointSize = 10;
        tableHead.leading = 16;
        tableHead.fillColor = 'Paper';

    }

    if (doc.pages[1].textFrames.item('eld_box').isValid && doc.pages[1].textFrames.item('eld_box').parentStory.tables[0].isValid) {

        var table = doc.pages[1].textFrames.item('eld_box').parentStory.tables[0];

        var cells = table.cells;
        var c = cells.length

        while (c--) {

            cells[c].topEdgeStrokeWeight = .5;
            cells[c].topEdgeStrokeTint = 100;
            cells[c].leftEdgeStrokeWeight = .5;
            cells[c].leftEdgeStrokeTint = 100;
            cells[c].rightEdgeStrokeWeight = .5;
            cells[c].rightEdgeStrokeTint = 100;
            cells[c].bottomEdgeStrokeWeight = .5;
            cells[c].bottomEdgeStrokeTint = 100;
            cells[c].fillColor = 'Paper';
            // cells[c].strokeColor = doc.swatches.item('Black');


            cells[c].width = 0.7451 + 'in';
            cells[c].height = 0.0417 + 'in';


            table.cells[0].fillColor = 'week1';
            table.cells[0].paragraphs[0].appliedParagraphStyle = 'ELD_table_head';

            table.cells[1].fillColor = 'week1';
            table.cells[1].paragraphs[0].appliedParagraphStyle = 'ELD_table_head';

            table.cells[2].fillColor = 'week1';
            table.cells[2].paragraphs[0].appliedParagraphStyle = 'ELD_table_head';

            app.findGrepPreferences.findWhat = '\\r(?=$)';
            app.changeGrepPreferences.changeTo = '';
            cells[c].changeGrep()
        }


    }

    if (doc.pages[3].textFrames.item('eld_box').isValid && doc.pages[3].textFrames.item('eld_box').parentStory.tables[0].isValid) {

        var table = doc.pages[3].textFrames.item('eld_box').parentStory.tables[0];

        var cells = table.cells;
        var c = cells.length


        while (c--) {

            cells[c].paragraphs[0].appliedParagraphStyle = 'ELD_table_body';

            cells[c].topEdgeStrokeWeight = .5;
            cells[c].topEdgeStrokeTint = 100;
            cells[c].leftEdgeStrokeWeight = .5;
            cells[c].leftEdgeStrokeTint = 100;
            cells[c].rightEdgeStrokeWeight = .5;
            cells[c].rightEdgeStrokeTint = 100;
            cells[c].bottomEdgeStrokeWeight = .5;
            cells[c].bottomEdgeStrokeTint = 100;
            // cells[c].strokeColor = doc.swatches.item('Black');


            cells[c].width = 0.7451 + 'in';
            cells[c].height = 0.0417 + 'in';


            table.cells[0].fillColor = 'week1';
            table.cells[0].paragraphs[0].appliedParagraphStyle = 'ELD_table_head';

            table.cells[1].fillColor = 'week1';
            table.cells[1].paragraphs[0].appliedParagraphStyle = 'ELD_table_head';

            table.cells[2].fillColor = 'week1';
            table.cells[2].paragraphs[0].appliedParagraphStyle = 'ELD_table_head';

            app.findGrepPreferences.findWhat = '\\r(?=$)';
            app.changeGrepPreferences.changeTo = '';
            cells[c].changeGrep()


        }


    }
}




function lesson_tabber(doc, session) {

    session = session.xmlContent.contents;
    session = session.replace(/\D/g, '');
    var tab = doc.textFrames.item('lesson_tab');
    var height = tab.geometricBounds[2] - tab.geometricBounds[0];
    var y = tab.geometricBounds[3];
    var x = tab.geometricBounds[0];
    var multiplier = 0;

    switch (session) {

        case '1':
        case '6':
        case '11':
        case '16':
        case '21':
        case '26':
            multiplier = 0;
            break;

        case '2':
        case '7':
        case '12':
        case '17':
        case '22':
        case '27':
            multiplier = 1;
            break;

        case '3':
        case '8':
        case '13':
        case '18':
        case '23':
        case '28':
            multiplier = 2;
            break;

        case '4':
        case '9':
        case '14':
        case '19':
        case '24':
        case '29':
            multiplier = 3;
            break;

        case '5':
        case '10':
        case '15':
        case '20':
        case '25':
        case '30':
            multiplier = 4;
            break;
    }

    x += (height * multiplier);

    if (x !== 0) {

        tab.move([y, x]);
    }


}

function meta(xml) {

    var xml = doc.xmlElements[0];
    var master1 = doc.masterSpreads.item('A-Lesson');
    var runningHdL = master1.pages[0].pageItems.item('running-hd-l');
    var runningHdR = master1.pages[1].pageItems.item('running-hd-r');

    var master3 = doc.masterSpreads.item('B-2-Box');
    var runningHdR3 = master3.pages[1].pageItems.item('running-hd-r');

    var master2 = doc.masterSpreads.item('C-3-Box');
    var runningHdL2 = master2.pages[0].pageItems.item('running-hd-l');
    var runningHdR2 = master2.pages[1].pageItems.item('running-hd-r');

    var navtitles = xml.evaluateXPathExpression('//navtitle');
    var grade = navtitles[0];
    var unit = navtitles[1];
    var week = navtitles[2];
    var session = navtitles[3];
    var genre = navtitles[4];
    var sessionGoal = navtitles[5];

    var lessonTab = doc.textFrames.item('lesson_tab');

    lesson_tabber(doc, session);
    lessonTab.contents = session.contents;

    try {
        app.findGrepPreferences.findWhat = '\\r(?=$)';
        app.changeGrepPreferences.changeTo = '';
        lessonTab.changeGrep();

    } catch (e) {}



    genre.contents.appliedCharacterStyle = 'meta_GENRE';

    runningHdL.contents = genre.contents + ' ' + week.contents + ': ' + session.contents;
    // runningHdL.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    // runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    // runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;
    replacer(runningHdL, '\\r', '', false, false, false, false, true);
    replacer(runningHdL, '(Part )(\\d)(\\s\\s*)', '$1$2\\r', false, false, false, false, true);
    runningHdL.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    replacer(runningHdL, '^(?:(?!Week).)*', '$0', false, false, false, 'meta_GENRE', true);
    runningHdL.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    replacer(runningHdL, 'Week\\s\\d\\s*:', '$0', false, false, false, 'meta_WEEK', true);
    runningHdL.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    replacer(runningHdL, 'Session\\s\\d\\d*', '$0', false, false, false, 'meta_SESSION', true);
    runningHdL.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    // replacer(runningHdL, 'Part\\s\\d', '$0', false, false, false, 'meta_PART', true);
    // runningHdL.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    // runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    // runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;


    // alert(runningHdL.paragraphs.length);
    // alert(runningHdL.contents + '\r' + runningHdL.contents.appliedCharacterStyle.name)

    runningHdR.contents = week.contents + ': ' + session.contents + '\r' + sessionGoal.contents;
    runningHdR3.contents = week.contents + ': ' + session.contents + '\r' + sessionGoal.contents;

    try {

        replacer(runningHdR, '(Week)(\\s)(\\d)(\\s*\\r)', '$1$2$3', false, false, false, false, true);
        replacer(runningHdR, '^\\r', '', false, false, false, false, true);
        replacer(runningHdR, 'Session\\s\\d\\d*', '$0', false, false, false, 'meta_SESSION', true);
        replacer(runningHdR, 'Week\\s\\d\\s*:', '$0', false, false, false, 'meta_WEEK', true);

        runningHdR.paragraphs[1].appliedCharacterStyle = doc.characterStyles.item('meta_SESSIONGOAL');
        runningHdR.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.item('Running-Hd_right');
        runningHdR.paragraphs[1].appliedParagraphStyle = doc.paragraphStyles.item('Running-Hd_right');

        replacer(runningHdR3, '(Week)(\\s)(\\d)(\\s*\\r)', '$1$2$3', false, false, false, false, true);
        replacer(runningHdR3, '^\\r', '', false, false, false, false, true);
        replacer(runningHdR3, 'Session\\s\\d\\d*', '$0', false, false, false, 'meta_SESSION', true);
        replacer(runningHdR3, 'Week\\s\\d\\s*:', '$0', false, false, false, 'meta_WEEK', true);

        runningHdR3.paragraphs[1].appliedCharacterStyle = doc.characterStyles.item('meta_SESSIONGOAL');
        runningHdR3.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.item('Running-Hd_right');
        runningHdR3.paragraphs[1].appliedParagraphStyle = doc.paragraphStyles.item('Running-Hd_right');

    } catch (e) {


    }

    runningHdL2.contents = genre.contents + ' ' + week.contents + ': ' + session.contents;
    runningHdL2.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;
    replacer(runningHdL2, '\\r', '', false, false, false, false, true);
    // replacer(runningHdL2, '(Part )(\\d)(\\s\\s*)', '$1$2\\r', false, false, false, false, true);
    runningHdL2.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    replacer(runningHdL2, '^(?:(?!Week).)*', '$0', false, false, false, 'meta_GENRE', true);
    runningHdL2.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    replacer(runningHdL2, 'Week\\s\\d\\s*:', '$0', false, false, false, 'meta_WEEK', true);
    runningHdL2.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    replacer(runningHdL2, 'Session\\s\\d\\d*', '$0', false, false, false, 'meta_SESSION', true);
    runningHdL2.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    // replacer(runningHdL2, 'Part\\s\\d', '$0', false, false, false, 'meta_PART', true);
    // runningHdL2.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    // runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    // runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;


    var page = xml.evaluateXPathExpression('//learningBECContent//page');
    doc.sections[0].pageNumberStart = Number(page[0].xmlAttributes.item('value').value);


    try {
        app.findGrepPreferences.findWhat = '\\r(?=$)';
        app.changeGrepPreferences.changeTo = '';
        runningHdL2.changeGrep();
        runningHdL.changeGrep();

    } catch (e) {}


    // runningHdL2.move([runningHdL2.geometricBounds[1], 0.3566415478176]);
    runningHdL.geometricBounds = [0.06666666666667 + 'in', 0.75 + 'in', 0.4 + 'in', 8.75 + 'in'];
    runningHdL2.geometricBounds = [0.06666666666667 + 'in', 0.75 + 'in', 0.4 + 'in', 8.75 + 'in'];
}

function clearOverrides(target) {

    var paras = target.paragraphs;
    var p = paras.length;

    while (p--) {

        paras[p].clearOverrides();
    }
}

function topic_placerLesson(xml, exceptions) {

    pacifier(xml);
    namer();

    var lesson = xml.evaluateXPathExpression('//learningBECContent[@outputclass="lesson"]//title');
    var lessonFrame = doc.pages[0].pageItems.item('lessonAhd');
    var lessonFrame2 = doc.pages[2].pageItems.item('lessonAhd');
    lesson[0].placeXML(lessonFrame);
    app.select(lessonFrame.paragraphs[0]);
    app.copy();
    app.select(lessonFrame2.insertionPoints[0]);
    app.paste();

}

function topic_placerTeach(xml, exceptions) {

    pacifier(xml);
    namer();

    var lesson = xml.evaluateXPathExpression('//learningBECContent[@outputclass="teach"]//title');
    var lessonFrame = doc.pages[1].pageItems.item('teach');
    lesson[0].placeXML(lessonFrame);
}


function topic_placerWW(xml, exceptions) {

    pacifier(xml);
    namer();

    // var topics = xml.evaluateXPathExpression('//learningBECContent[@outputclass="teach"]');
    // var t = topics.length;

    var sections = xml.evaluateXPathExpression('//learningBECContent[@outputclass="teach"]//section|//learningBECContent[@outputclass="lesson"]//section');
    var s = sections.length;

    // alert(sections[0].xmlContent.contents);
    // exit();




    // while (s--) {

    // for (p = 0; p < doc.pages.length; p++) {

    for (a = 0; a < sections.length; a++) {

        var sectionFrame = doc.pageItems.item('section' + '_' + a);

        // for (b = 0; b < sectionFrame.length; b++) {

        sections[a].placeXML(sectionFrame);

        // }
        // }
    }
}

function topic_placerW1(xml, exceptions) {

    pacifier(xml);
    namer();

    // var topics = xml.evaluateXPathExpression('//learningBECContent[@outputclass="teach"]');
    // var t = topics.length;

    var sections = xml.evaluateXPathExpression('//learningBECContent[@outputclass="lesson"]//section');
    var s = sections.length;

    for (a = 0; a < sections.length; a++) {

        var sectionFrame = doc.pageItems.item('lesson' + a);
        sections[a].placeXML(sectionFrame);
        // alert(a);
        // exit();
    }

    var lesson4 = doc.pageItems.item('lesson4');
    var lesson5 = doc.pageItems.item('lesson5');
    app.select(lesson5.texts.everyItem());
    app.copy();
    lesson5.remove();
    app.select(lesson4.insertionPoints[-2]);
    app.paste();
    // exit();


}


function topic_placer(xml, exceptions) {

    pacifier(xml);
    namer();

    var topics = xml.evaluateXPathExpression('//learningBECContent');
    var t = topics.length;

    while (t--) {

        var place = true;
        pacifier(topics[t]);
        var type = topics[t].xmlAttributes.item('outputclass').value;
        var e = exceptions.length;
        pacifier(topics[t]);

        while (e--) {

            if (exceptions[e] == type) {

                place = false;
            }
        }

        if (place) {

            var page = topics[t].evaluateXPathExpression('//page');

            try {

                var p = page[0].xmlAttributes.item('value').value.replace(/\D/g, '');

                if (p.match(/^00*/)) {

                    p = p.replace(/^00*/, '');
                }

            } catch (e) { alert('NO PAGE NUMBER! topic type: ' + type + ' ' + e.line + ': ' + e) }

            if (!app.activeDocument.pages.itemByName(p.toString()).isValid) {

                alert('PAGE DOESNT EXIST');
                exit();
            }

            var frames = getTextFrames(app.activeDocument.pages.itemByName(p.toString()));
            var f = frames.length;

            while (f--) {

                if (frames[f].name == type && frames[f].paragraphs.length == 0) {

                    frame = frames[f];
                    topics[t].placeXML(frame);
                }
            }
        }
    }
}

function cleanup(doc) {


    var eldLesson = doc.pages[1].pageItems.item('eld_box');
    var eldLesson2 = doc.pages[3].pageItems.item('eld_box');

    eldLesson.itemLayer = doc.layers.item('eld_article');
    eldLesson2.itemLayer = doc.layers.item('eld_article');



    replacer(doc, '\\s$', '', false, false, false, false, false);
    replacer(doc, '\\s$', '', false, false, false, false, false);

    if (doc.pageItems.item('eld_group').isValid) {

        doc.pageItems.item('eld_group').remove();
    }

    var frames = doc.textFrames;
    var f = frames.length;

    if (!doc.name.match(/w1|w6/i)) {

        var lessonBhd = doc.paragraphStyles.item('lesson_B-hd');
        lessonBhd.appliedFont = 'ITC Avant Garde Gothic Std';
        lessonBhd.fontStyle = 'Bold';
        lessonBhd.pointSize = 12;
        lessonBhd.leading = 14;

        replacer(doc, 'THE GOALS:', '$0', 'goals_B-hd', false, 'goals_B-hd_noSpaceBefore', 'red', false);
        replacer(doc, 'Independent Writing and Feedback', 'Independent Writing\nand Feedback', 'lesson_B-hd', false, false, false, false);
        replacer(doc, '(Teach the Strategy)(\\s)(\\(.*\\))', '$1\\n$3', 'teach_A-hd', false, false, false, false);

    } else {

        replacer(doc, 'State the Focus and Purpose', '$0', 'lesson_A-hd', false, 'lesson_B-hd', false, false);
        replacer(doc, 'Compare Elaboration Techniques', '$0', 'lesson_A-hd', false, 'lesson_B-hd', false, false);
        replacer(doc, 'Share and Reflect', '$0', 'lesson_A-hd', false, 'lesson_B-hd', false, false);


    }

    if (doc.pages[0].pageItems.item('section_0').isValid) {

        var section0 = doc.pages[0].pageItems.item('section_0');
        section0.paragraphs[0].appliedParagraphStyle = lessonBhd;
    }






    try {

        while (f--) {

            app.findGrepPreferences.findWhat = '\\r(?=$)';
            app.changeGrepPreferences.changeTo = '';
            frames[f].changeGrep();
            frames[f].fit(FitOptions.FRAME_TO_CONTENT);

        }


    } catch (e) {}


    try {
        app.findGrepPreferences.findWhat = '\\r(?=$)';
        app.changeGrepPreferences.changeTo = '';
        frames[f].changeGrep();
        frames[f].fit(FitOptions.FRAME_TO_CONTENT);

    } catch (e) {}

    var keep = doc.paragraphStyles.length;

    while (keep--) {

        if (keep !== 0) {
            doc.paragraphStyles[keep].hyphenation = false;
        }
    }

    if (doc.name.match(/w1|w6/i)) {


        try {

            while (f--) {

                app.findGrepPreferences.findWhat = '\\r(?=$)';
                app.changeGrepPreferences.changeTo = '';
                frames[f].changeGrep();
                frames[f].fit(FitOptions.FRAME_TO_CONTENT);

            }


        } catch (e) {}

        var strat = doc.pages[2].pageItems.item('strategy');
        var lesson0 = doc.pageItems.item('lesson0');
        var lesson1 = doc.pageItems.item('lesson1');
        var lesson2 = doc.pageItems.item('lesson2');
        var lesson3 = doc.pages[2].pageItems.item('lesson3');
        var lesson4 = doc.pages[2].pageItems.item('lesson4');
        var lesson5 = doc.pages[2].pageItems.item('lesson5');
        var lesson6 = doc.pages[2].pageItems.item('lesson6');
        var goals = doc.pages[0].pageItems.item('goals');
        var pdLessonGroup = doc.pages[1].pageItems.item('PD_box');


        goals.move([goals.geometricBounds[1], lesson0.geometricBounds[2] + 0.25]);
        lesson1.move([lesson1.geometricBounds[1], goals.geometricBounds[2] + 0.25]);
        strat.move([lesson3.geometricBounds[1], lesson3.geometricBounds[2] + 0.25]);
        lesson4.move([strat.geometricBounds[1], strat.geometricBounds[2] + 0.25]);

        pdLessonGroup.move([lesson2.geometricBounds[1], lesson2.geometricBounds[2] + 0.35]);


    } else {


        // time cleanup

        try {

            // var section4 = doc.pageItems.item('section4');
            // var section6 = doc.pageItems.item('section6');

            replacer(doc, '\\(\\d.*\\)', '\n$0', 'lesson_B-hd', false, false, false, false);
            replacer(doc, '\\(\\d.*\\)', '\n$0', 'lesson_B-hd', false, false, false, false);
            replacer(doc, '\\n\\n', '\\n', false, false, false, false, false)

        } catch (e) {

        }

        // Spread 1 pageitems 

        var goals = doc.pages[0].pageItems.item('goals');
        var materials = doc.pages[0].pageItems.item('materials');
        var strat = doc.pages[0].pageItems.item('strategy');
        var eld1Icon = doc.pages[1].pageItems.item('el_icon');
        var eld1Frame = doc.pages[1].pageItems.item('eld_box');
        var eldArray = [];
        var pdLessonGroup = doc.pages[0].pageItems.item('PD_box');
        var FPO = doc.pages[1].pageItems.item('FPO');

        // Spread 2 pageitems

        var lookFors = doc.pageItems.item('look_fors');
        var strat2 = doc.pages[2].pageItems.item('strategy');
        var strat3 = doc.pages[3].pageItems.item('strategy');
        var edit = doc.pages[2].pageItems.item('edit');
        var pdLessonGroup2 = doc.pages[3].pageItems.item('PD_box');
        var eld2Icon = doc.pages[3].pageItems.item('el_icon');
        var eld2Frame = doc.pages[3].pageItems.item('eld_box');
        var eldArray2 = [];

        // Spread 2 lesson body text frames

        var lesson0 = doc.pages[2].pageItems.item('lesson');
        var bridgeLesson = doc.pages[2].pageItems.item('bridge');
        var lesson1 = doc.pages[2].pageItems.item('lesson1');
        var lesson2 = doc.pages[3].pageItems.item('lesson2');
        var lesson3 = doc.pages[3].pageItems.item('lesson3');

        // Spread 1 lesson body text frames

        var col1 = doc.pages[0].pageItems.item('1_col1');
        var col2 = doc.pages[0].pageItems.item('1_col2');
        var col3 = doc.pages[1].pageItems.item('1_col3');
        var col4 = doc.pages[1].pageItems.item('1_col4');

        // fit spread2 frames to content

        var frames = doc.textFrames;
        var f = frames.length;

        try {
            lesson0.fit(FitOptions.FRAME_TO_CONTENT);
            lesson1.fit(FitOptions.FRAME_TO_CONTENT);
            lesson2.fit(FitOptions.FRAME_TO_CONTENT);
            lesson3.fit(FitOptions.FRAME_TO_CONTENT);
            edit.fit(FitOptions.FRAME_TO_CONTENT);
            bridgeLesson.fit(FitOptions.FRAME_TO_CONTENT);

        } catch (e) {

            alert('Error closing up textframes.\n' + e);
        }

        // Add spacing between objects
        col1.move([0.8738, 2.35]);
        col2.move([4.735, col1.geometricBounds[0]]);
        goals.move([col1.geometricBounds[1], col1.geometricBounds[2] + 0.25]);
        materials.move([goals.geometricBounds[1], materials.geometricBounds[0]]);
        strat.move([col2.geometricBounds[1], col2.geometricBounds[2] + 0.25]);
        pdLessonGroup.move([strat.geometricBounds[1], strat.geometricBounds[2] - 0.032]);
        col3.move([11, col2.geometricBounds[0]]);
        col4.move([14.75, col3.geometricBounds[0]]);
        FPO.move([10.0706, col3.geometricBounds[2] + 0.25]);
        eld1Frame.move([col4.geometricBounds[1], col4.geometricBounds[2] + 0.25]);

        // eldArray.push(eld1Icon);
        // eldArray.push(eld1Frame);
        // var eld1Group = doc.groups.add(eldArray);
        // eld1Group.move([14.8199, col4.geometricBounds[2] - 0.1]);
        // eld1Group.ungroup();

        // editArray.push(editIcon);
        // editArray.push(edit);
        // var editGroup = doc.groups.add(editArray);

        lesson0.move([0.9, 2.3667]);
        lookFors.move([lesson0.geometricBounds[1], lesson0.geometricBounds[2] + 0.25]);
        bridgeLesson.move([lesson0.geometricBounds[1], lookFors.geometricBounds[2] + 0.25]);
        lesson1.move([4.735, lesson0.geometricBounds[0]]);
        strat2.move([lesson1.geometricBounds[1], lesson1.geometricBounds[2] + 0.25]);
        // editGroup.move([strat2.geometricBounds[1], strat2.geometricBounds[2] + 0.25]);
        // editGroup.ungroup();
        lesson2.move([11, lesson1.geometricBounds[0]]);
        strat3.move([lesson2.geometricBounds[1], lesson2.geometricBounds[2] + 0.25]);
        pdLessonGroup2.move([lesson2.geometricBounds[1], strat3.geometricBounds[2] - 0.032]);
        lesson3.move([14.85, lesson2.geometricBounds[0]]);
        eld2Frame.move([lesson3.geometricBounds[1], lesson3.geometricBounds[2] + 0.25]);

        // eldArray2.push(eld2Icon);
        // eldArray2.push(eld2Frame);
        // var eld2Group = doc.groups.add(eldArray2);
        // eld2Group.move([lesson3.geometricBounds[1], lesson3.geometricBounds[2] - 0.1]);
        // eld2Group.ungroup();

        iconFix(doc);



    }



    // var pdFrame = doc.pages[0].textFrames.item('pd');
    // var pdFrame2 = doc.pages[3].textFrames.item('pd');

    // if (pdFrame.paragraphs.length == 0) {

    //     var pdBox = doc.pages[0].pageItems.item('PD_box');
    //     pdFrame.remove();
    //     pdBox.remove();
    // }

    // if (pdFrame2.paragraphs.length == 0) {

    //     var pdBox = doc.pages[3].pageItems.item('PD_box');
    //     pdFrame2.remove();
    //     pdBox.remove();
    // }


}

function iconFix(doc) {

    var eld1Icon = doc.pages[1].pageItems.item('el_icon');
    var eld1Frame = doc.pages[1].pageItems.item('eld_box');
    var eld2Icon = doc.pages[3].pageItems.item('el_icon');
    var eld2Frame = doc.pages[3].pageItems.item('eld_box');
    try {
        var edit = doc.pages[2].pageItems.item('edit');
        var editIcon = doc.pages[2].pageItems.item('two_minute_icon');
    } catch (e) {

        var gramGroup = doc.pageItems.item('grammar_group');
        gramGroup.ungroup();
    }

    var eldYSpace = 0.0958;
    var eldXSpace = 0.1111;
    var editXSpace = 0.1677;
    var editYSpace = -0.08;

    eld1Icon.move([eld1Frame.geometricBounds[1] - eldXSpace, eld1Frame.geometricBounds[0] - eldYSpace]);
    eld2Icon.move([eld2Frame.geometricBounds[1] - eldXSpace, eld2Frame.geometricBounds[0] - eldYSpace]);
    try {
        editIcon.move([edit.geometricBounds[1] - editXSpace, edit.geometricBounds[0] - editYSpace]);
    } catch (e) {

    }

    // icon 2.98072187466612,14.8338731476707,3.41306196943187,15.2644347576755
    // frame 3.06947310096353,14.9379193943706,8.14555766178167,18.055670378772


}

function metaStyles(doc) {

    // Add metadata char styles

    try {

        var genre = doc.characterStyles.add();
        genre.name = 'meta_GENRE';
        genre.appliedFont = 'ITC Avant Garde Gothic Std';
        genre.fontStyle = 'Bold';
        genre.pointSize = 12;
        genre.leading = 16;

    } catch (e) {

        // alert(e);

    }

    try {

        var week = doc.characterStyles.add();
        week.name = 'meta_WEEK';
        week.appliedFont = 'ITC Avant Garde Gothic Std';
        week.fontStyle = 'Book';
        week.pointSize = 12;
        week.leading = 16;

    } catch (e) {

        // alert(e);

    }

    try {

        var session = doc.characterStyles.add();
        session.name = 'meta_SESSION';
        session.appliedFont = 'ITC Avant Garde Gothic Std';
        session.fontStyle = 'Bold';
        session.pointSize = 12;
        session.leading = 16;

    } catch (e) {

        // alert(e);

    }

    try {

        var part = doc.characterStyles.add();
        part.name = 'meta_PART';
        part.appliedFont = 'Handwriting';
        part.fontStyle = 'Regular';
        part.pointSize = 33;
        part.leading = 16;

    } catch (e) {

        // alert(e);

    }

    try {

        var sessiontitle = doc.paragraphStyles.add();
        sessiontitle.name = 'sessiontitle_A-hd';
        sessiontitle.appliedFont = 'Rockwell';
        sessiontitle.fontStyle = 'Light';
        sessiontitle.pointSize = 20;
        sessiontitle.leading = 25;

    } catch (e) {

        // alert(e);

    }
}

function elLessonPlacer(xml) {

    var elLesson = xml.evaluateXPathExpression('//learningBECContent[@outputclass="el_lesson"]');
    var elLessonFrame = doc.pages[1].pageItems.item('eld_box');
    var elLessonGroup = doc.pages[1].pageItems.item('eld_group');
    var master1 = doc.masterSpreads.item('A-Lesson');

    if (elLesson.length == 1) {

        // alert('1');
        elLessonGroup.ungroup();
        elLesson[0].placeXML(elLessonFrame);


    } else if (elLesson.length == 2) {

        // alert('2');

        var elLessonFrame = doc.pages[1].pageItems.item('eld_box');
        var elLessonGroup = doc.pages[1].pageItems.item('eld_group');
        var elLessonFrame2 = doc.pages[3].pageItems.item('eld_box');
        var elLessonGroup2 = doc.pages[3].pageItems.item('eld_group')

        elLessonGroup.ungroup();
        elLessonGroup2.ungroup();
        elLesson[1].placeXML(elLessonFrame2);
        elLesson[0].placeXML(elLessonFrame);

    } else {

        try {

            elLessonGroup.remove();
            elLessonGroup2.remove();

        } catch (e) {


        }
    }


    // alert(elLesson[0].xmlContent.contents);

}

function pdPlacer(xml) {

    var pdLesson = xml.evaluateXPathExpression('//learningBECContent[@outputclass="pd"]');
    var pdLessonFrame = doc.pages[0].pageItems.item('pd');
    var pdLessonGroup = doc.pages[0].pageItems.item('PD_box');
    // var master1 = doc.masterSpreads.item('A-Lesson');

    if (pdLesson.length == 1) {

        // alert('1');
        // alert(pdLesson[0].contents);
        try {
            pdLessonGroup.ungroup();
            pdLesson[0].placeXML(pdLessonFrame);
        } catch (e) {
            pdLessonGroup = doc.pages[3].pageItems.item('PD_box');
            pdLessonFrame = doc.pages[3].pageItems.item('pd');
            pdLessonGroup.ungroup();
            pdLesson[0].placeXML(pdLessonFrame);
        }



    } else if (pdLesson.length == 2) {

        // alert('2');

        var pdLessonFrame = doc.pages[0].pageItems.item('pd');
        var pdLessonGroup = doc.pages[0].pageItems.item('PD_box');
        var pdLessonFrame2 = doc.pages[3].pageItems.item('pd');
        var pdLessonGroup2 = doc.pages[3].pageItems.item('PD_box')

        pdLessonGroup.ungroup();
        pdLessonGroup2.ungroup();
        pdLesson[1].placeXML(pdLessonFrame2);
        pdLesson[0].placeXML(pdLessonFrame);

    } else {

        try {

            pdLessonGroup.remove();
            pdLessonGroup2.remove();

        } catch (e) {


        }
    }

    if (doc.pages[0].pageItems.item('PD_box').isValid) {

        doc.pages[0].pageItems.item('PD_box').remove();
    }

    if (doc.pages[1].pageItems.item('PD_box').isValid) {

        doc.pages[1].pageItems.item('PD_box').remove();
    }

    if (doc.pages[2].pageItems.item('PD_box').isValid) {

        doc.pages[2].pageItems.item('PD_box').remove();
    }

    if (doc.pages[3].pageItems.item('PD_box').isValid) {

        doc.pages[3].pageItems.item('PD_box').remove();
    }


    // alert(elLesson[0].xmlContent.contents);

}

// function pdPlacer(xml) {

//     var pdLesson = xml.evaluateXPathExpression('//learningBECContent[@outputclass="pd"]');
//     var pdFrame = doc.pages[0].pageItems.item('quote');
//     var pdFrame2 = doc.pages[3].pageItems.item('quote');
//     var master = doc.masterSpreads.item('D-Icons');
//     // var pdLessonGroup = doc.pages[0].pageItems.item('pd_group');
//     // var pdLessonGroup2 = doc.pages[3].pageItems.item('pd_group');

//     if (pdLesson.length > 0 && pdLesson.length < 3) {

//         for (x = 0; x < pdLesson.length; x++) {

//             var page = pdLesson[x].evaluateXPathExpression('//page');
//             var p = page[0].xmlAttributes.item('value').value.replace(/\D/g, '');

//             if (p) {}
//         }

//         pdLessonGroup.ungroup();
//         // pdLessonGroup2.ungroup();

//         pdLesson[0].placeXML(pdFrame);
//         // pdLesson[1].placeXML(pdFrame2);

//     } else {

//         alert('NO PD LESSON');
//     }



// }
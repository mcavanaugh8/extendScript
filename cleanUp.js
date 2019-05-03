/* beautify preserve:start */
#include '/Volumes/BEC/15 User Files/INDD SCRIPTS/primitive/main.js';
/* beautify preserve:end */

var doc = app.activeDocument;

var input = prompt('Actions Menu:\n1. Cleanup\n2. Icon Fix\n3. Part 1 & 2 Metadata\n4. Thumbnail Placer\n\nType the number of the process you wish to run.', '');


if (input == '1' || input == 'Cleanup' || input == 'cleanup') {

    if (doc.name.match(/w2|w3|w4|w5/i)) {

        cleanup(doc);

    } else if (doc.name.match(/w1|w6/)) {

        alert('Week 1 & Week 6');
        exit();
    }


} else if (input == '2' || input == 'Icon Fix' || input == 'icon fix') {

    iconFix(doc);

} else if (input == '3') {

	if (doc.name.match(/w2|w3|w4|w5/i)) {

		alert('This file is not Week 1 or Session 29 or 30. It does not require Part 1 or Part 2 in the header.');
		exit();

	} else if (doc.name.match(/w1||s29||s30/i)) {

    var master1 = doc.masterSpreads.item('A-Lesson');
    var master2 = doc.masterSpreads.item('C-3-Box');
    var master3 = doc.masterSpreads.item('B-2-Box');

    var runningHdL = master1.pages[0].pageItems.item('running-hd-l');
    var runningHdL2 = master2.pages[0].pageItems.item('running-hd-l');
    var runningHdR3 = master3.pages[1].pageItems.item('running-hd-r');

    replacer(runningHdL, 'Session\\s\\d\\d*', '$0\\s\\sPart 1', false, false, false, 'meta_SESSION', true);
    runningHdL.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    replacer(runningHdL, 'Part\\s\\d', '$0', false, false, false, 'meta_PART', true);
    runningHdL.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    replacer(runningHdL2, 'Session\\s\\d\\d*', '$0\\s\\sPart\\s2', false, false, false, 'meta_SESSION', true);
    runningHdL2.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;

    replacer(runningHdL2, 'Part\\s\\d', '$0', false, false, false, 'meta_PART', true);
    runningHdL2.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_CENTER_POINT;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    runningHdL2.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;
}

} else if (input == '4') {

    alert('Not done yet :)');
    exit();
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

function cleanup(doc) {

    // iconFix(doc);

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
    var editIcon = doc.pages[2].pageItems.item('two_minute_icon');
    var pdLessonGroup2 = doc.pages[3].pageItems.item('PD_box');
    var eld2Icon = doc.pages[3].pageItems.item('el_icon');
    var eld2Frame = doc.pages[3].pageItems.item('eld_box');
    var eldArray2 = [];
    var editArray = [];

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
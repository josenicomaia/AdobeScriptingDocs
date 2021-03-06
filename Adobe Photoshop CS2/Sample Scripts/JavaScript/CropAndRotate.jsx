// Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
// Crop and rotate the active document.

// enable double clicking from the Macintosh Finder or the Windows Explorer
#target photoshop

// in case we double clicked the file
app.bringToFront();

// debug level: 0-2 (0:disable, 1:break on error, 2:break at beginning)
// $.level = 0;
// debugger; // launch debugger on next line

if (!app.documents.length > 0) {    // open sample file if no document is opened.
    var fileName = app.path.toString() + "/Samples/Dune.tif";
    var docRef = open( File(fileName) );
}

app.preferences.rulerUnits = Units.PIXELS;
// crop a 10 pixel border from the image
bounds = new Array(10, 10, app.activeDocument.width - 10, app.activeDocument.height - 10);
app.activeDocument.rotateCanvas(45);
app.activeDocument.crop(bounds);
bounds = null;

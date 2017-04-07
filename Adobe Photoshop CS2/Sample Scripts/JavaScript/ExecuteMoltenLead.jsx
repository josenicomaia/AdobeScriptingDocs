// Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.

// This script demonstrates how to use the action manager to execute a 
// previously defined action. The name of the action comes from
// Photoshop's Actions Palette and may be different if running a non-English version of Photoshop

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

try {
	doAction("Molten Lead", "Sample Actions.atn");
} catch (Error) {
	alert("Please load \"Sample Actions\" actions set.");
}

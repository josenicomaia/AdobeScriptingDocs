// Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
// Open a Photoshop document located in the Photoshop samples folder on the Photoshop CD.
// You must first create a File object to pass into the open method.

// enable double clicking from the Macintosh Finder or the Windows Explorer
#target photoshop

// in case we double clicked the file
app.bringToFront();

// debug level: 0-2 (0:disable, 1:break on error, 2:break at beginning)
// $.level = 0;
// debugger; // launch debugger on next line

var fileRef = new File(app.path.toString() + "/Samples/Dune.tif");
open (fileRef);
fileRef = null;



' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' This script executes a JavaScript. The script MosaicTiles.js is located 
' in the JavaScript folder.

Option Explicit

Dim appRef
Dim javaScriptFile

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

javaScriptFile = appRef.Path & "Scripting Guide\Sample Scripts\JavaScript\MosaicTiles.jsx"

appRef.DoJavaScriptFile( javaScriptFile )

MsgBox "Execute JavaScript complete"

' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' This script executes a JavaScript. The script MosaicTiles.js is located 
' in the JavaScript folder.

Option Explicit
On Error Resume Next

Dim appRef
Dim javaScriptFile

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

javaScriptFile = appRef.Path & "Scripting\Sample Scripts\JavaScript\MosaicTiles.jsx"

appRef.DoJavaScriptFile( javaScriptFile )

MsgBox "Execute JavaScript complete"

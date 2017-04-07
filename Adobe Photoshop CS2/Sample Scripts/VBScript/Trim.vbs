' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' This script demonstrates how to trim either the left and right edges of a document
' or the top and bottom edges of the document.

Option Explicit

Dim appRef
Dim docRef
Dim fileName

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

If appRef.Documents.Count > 0 Then
    Set docRef = appRef.ActiveDocument
Else ' open sample file
    fileName = appRef.Path & "\Samples\Ducky.tif"
    Set docRef = appRef.Open( fileName )
End If

Set docRef = appRef.ActiveDocument

docRef.Trim 1, False, True, False, True ' 1 = psTopLeftPixel

docRef.Trim 1, True, False, True, False

MsgBox "Trim complete"

' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' This script will create a new target document which is 4 inch x 4 inch document,
' select the contents of the source document and copy it to the clipboard,
' and then paste the contents of the clipboard into the target document.
' Notice that the script sets the active document prior to doing the cut
' and paste because these operations only work on the active document.

Option Explicit

Dim appRef
Dim docRef
Dim docRef2
Dim newLayerRef
Dim fileName

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

' now create a new document and copy and paste.
If appRef.Documents.Count > 0 Then
    Set docRef = appRef.ActiveDocument
Else ' open sample file
    fileName = appRef.Path & "\Samples\Dune.tif"
    Set docRef = appRef.Open( fileName )
End If

appRef.Preferences.RulerUnits = 2 ' psInches

Set docRef2 = appRef.Documents.Add( 4, 4, 72, "The New Document" )

appRef.ActiveDocument = docRef

docRef.Selection.SelectAll

docRef.Selection.Copy

appRef.ActiveDocument = docRef2

Set newLayerRef = docRef2.Paste

MsgBox "Apply Style script complete"

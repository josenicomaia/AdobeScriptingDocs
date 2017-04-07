' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' Apply the "Puzzle (Image)" layer style to the active layer

Option Explicit

Dim appRef
Dim docRef
Dim fileName
Dim artLayerRef

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

If appRef.Documents.Count > 0 Then
    Set docRef = appRef.ActiveDocument
Else ' open sample file
    fileName = appRef.Path & "\Samples\Dune.tif"
    Set docRef = appRef.Open( fileName )
End If

Set artLayerRef = docRef.ActiveLayer

If artLayerRef.IsBackgroundLayer Then
	artLayerRef.IsBackgroundLayer = False
End If
    
docRef.ActiveLayer.ApplyStyle "Puzzle (Image)"

MsgBox "Apply Style script complete"

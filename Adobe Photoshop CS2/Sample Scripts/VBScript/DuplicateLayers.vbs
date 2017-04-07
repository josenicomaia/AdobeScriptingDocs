' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' This script demonstrates how to add a LayerSet
' and then duplicate the top layer and place it into
' the Layer Set.

Option Explicit

Dim appRef
Dim docRef
Dim layerSetRef
Dim layerRef
Dim fileName

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

If appRef.Documents.Count > 0 Then
    Set docRef = appRef.ActiveDocument
Else ' open sample file
    fileName = appRef.Path & "\Samples\Dune.tif"
    Set docRef = appRef.Open( fileName )
End If

Set layerSetRef = docRef.LayerSets.Add
Set layerRef = docRef.ArtLayers(1).Duplicate
layerRef.MoveToEnd layerSetRef

MsgBox "Duplicate Layers complete"

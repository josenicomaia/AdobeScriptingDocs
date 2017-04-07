' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' Hello World Script

Option Explicit

Dim appRef
Dim strtRulerUnits
Dim strtTypeUnits
Dim textColor
Dim docRef
Dim artLayerRef
Dim layerKind
Dim textItemRef

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

strtRulerUnits = appRef.Preferences.RulerUnits
appRef.Preferences.RulerUnits = 2 ' psInches
strtTypeUnits = appRef.Preferences.TypeUnits
appRef.Preferences.TypeUnits = 5 ' psTypePoints

Set textColor = CreateObject( "Photoshop.SolidColor" )
textColor.RGB.Red = 255
textColor.RGB.Green = 0
textColor.RGB.Blue = 0

' Create a new document and assign it to a variable.
Set docRef = appRef.Documents.Add( 7, 5 )

'Create a new art layer, set it to a text layer.
Set artLayerRef = docRef.ArtLayers.Add

layerKind = 2 ' psTextLayer
artLayerRef.Kind = layerKind

' Set the contents and other properties of the text layer.
Set textItemRef = artLayerRef.TextItem
textItemRef.Contents = "Hello, World!"
textItemRef.Position = Array( 0.75, 0.75 )
textItemRef.Size = 36
textItemRef.Font = "Georgia"
textItemRef.Color = textColor

appRef.Preferences.RulerUnits = strtRulerUnits
appRef.Preferences.TypeUnits = strtTypeUnits

MsgBox "Create New Text Art complete"

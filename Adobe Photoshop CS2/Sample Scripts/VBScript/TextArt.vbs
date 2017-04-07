' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' This script shows how to determine if a layer is a text layer and
' how to apply to filter to a selection on the layer.
' The text layer must be rasterized before applying filters to it.
' Before running the script, create one or more text layers in the
' active document

Option Explicit

Dim appRef
Dim docRef
Dim textItemRef
Dim artLayerRef
Dim theOrigin
Dim theRasterizeType
Dim selectionType
Dim rippleSize
Dim boxWidth
Dim boxHeight
Dim layerKind
Dim textColor
Dim strtRulerUnits
Dim strtTypeUnits

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

If appRef.Documents.Count > 0 Then
    Set docRef = appRef.ActiveDocument
Else ' open new document with text
  'Create a new document and assign it to a variable.
  Set docRef = appRef.Documents.Add( 7, 5 )
End If

'Create a new art layer, set it to a text layer.
Set artLayerRef = docRef.ArtLayers.Add

layerKind = 2 ' psTextLayer
artLayerRef.Kind = layerKind

'Set the contents and other properties of the text layer.
Set textItemRef = artLayerRef.TextItem
textItemRef.Contents = "Hello, World!"
textItemRef.Position = Array( 0.75, 0.75 )
textItemRef.Size = 36
textItemRef.Font = "Georgia"
textItemRef.Color = textColor

theRasterizeType = 5 ' psEntireLayer
selectionType =  1 ' psReplaceSelection
rippleSize = 3 ' psLargeRipple

For Each artLayerRef In docRef.ArtLayers
    If ( artLayerRef.Kind = 2 ) Then ' psTextLayer
        docRef.ActiveLayer = artLayerRef
        
        ' must set the text kind to paragraph because you can only get the bounds of paragraph text.
        artLayerRef.TextItem.Kind = 2 ' psParagraphText
                    
        theOrigin = artLayerRef.TextItem.Position
        boxWidth = artLayerRef.TextItem.Width
        boxHeight = artLayerRef.TextItem.Height
        
        'Select the Text Art.
        'The origin takes the justification into account, so no need to test
        'if right/left/center justified.
        docRef.Selection.Select Array( Array( theOrigin( 0 ), theOrigin( 1 ) ), Array( boxWidth + theOrigin( 0 ), theOrigin( 1 ) ), Array( boxWidth + theOrigin( 0 ), theOrigin( 1 ) + boxHeight ), Array( theOrigin( 0 ), theOrigin( 1 ) + boxHeight ) ), selectionType, 0, False
        
        ' must rasterisze text items before applying filters to them.
        artLayerRef.Rasterize theRasterizeType
        artLayerRef.ApplyRipple 150, rippleSize
        
        artLayerRef.ApplyStyle ( "Overspray (Text)" )
        docRef.Selection.Deselect
    End If
Next

appRef.Preferences.RulerUnits = strtRulerUnits
appRef.Preferences.TypeUnits = strtTypeUnits
    
MsgBox "Text Art complete"

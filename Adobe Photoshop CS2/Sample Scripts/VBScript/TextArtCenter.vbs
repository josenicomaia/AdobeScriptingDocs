' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' Calculate the geometric center of each text item of a document

Option Explicit

Dim appRef
Dim selectedObjects
Dim objectBounds
Dim objectCenter
Dim textItemRef
Dim artLayerRef
Dim theTextType
Dim docRef
Dim layerKind
Dim textColor
Dim strtRulerUnits
Dim strtTypeUnits

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

theTextType = 2 ' psParagraphText

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

For Each artLayerRef In docRef.ArtLayers
    If ( artLayerRef.Kind = 2 ) Then ' psTextLayer

		docRef.ActiveLayer = artLayerRef
            
        ' must set the text kind to paragraph because you can only get the bounds of paragraph text.
        artLayerRef.TextItem.Kind = theTextType
        objectBounds = Array( artLayerRef.TextItem.Position( 0 )( 0 ), artLayerRef.TextItem.Position( 0 )( 1 ), artLayerRef.TextItem.Width, artLayerRef.TextItem.Height )
        objectCenter = GetItemCenter( objectBounds )
        MsgBox "Center of Text Item  x: " & objectCenter( 0 ) & ",  y :" & objectCenter( 1  )
    End If
Next

' The following lines define the function
Function GetItemCenter( ByVal sourceBounds )
	Dim left
	Dim top
	Dim right
	Dim bottom
	Dim xCenter
	Dim yCenter
	left = sourceBounds( 0 )
	top = sourceBounds( 1 )
	right = sourceBounds( 2 )
	bottom = sourceBounds( 3 )
	xCenter = ( left + right ) / 2
	yCenter = ( top + bottom ) / 2
	GetItemCenter = Array( xCenter, yCenter )
End Function

appRef.Preferences.RulerUnits = strtRulerUnits
appRef.Preferences.TypeUnits = strtTypeUnits

MsgBox "Text Art Center complete"

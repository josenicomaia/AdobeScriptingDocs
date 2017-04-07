' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' This script will iterate through all of the layers of a document. If the
' layer is not a text layer, the layer's color values are inverted. If it 
' is the background layer, also rotate, then crop the entire canvas. In 
' order to invert the entire document, each layer must be inverted 
' independently (or flattened beforehand).
'
' Before running this script, create a document with a few non-text layers.

Option Explicit

Dim appRef
Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

Dim docRef
Dim cropBounds
Dim artLayerRef
Dim strtRulerUnits

If ( appRef.Documents.Count > 0 ) Then
    strtRulerUnits = appRef.Preferences.RulerUnits
    appRef.Preferences.RulerUnits = 1 ' psPixels
    
    Set docRef = appRef.ActiveDocument
    Dim offset
    offset = 20
    cropBounds = Array( 20, 20, docRef.Width - offset, docRef.Height - offset )
        
    ' Check each ArtLayer to see what type it is.  Ignore all text layers.
    For Each artLayerRef In docRef.ArtLayers
    
        ' For every non-text layer, invert the contents.
        If ( Not artLayerRef.Kind = 2 ) Then ' psTextLayer
        
            ' Need to make the active layer this non-text layer in order to modify the layer.
            ' I want to invert the contents of the entire doc (excluding text), so need to
            ' invert every layer.
            docRef.ActiveLayer = artLayerRef
            
            ' See if there is anything on this layer, you can't invert an empty layer
            Dim w, h
            w = artLayerRef.Bounds(0)(2) - artLayerRef.Bounds(0)(0)
            h = artLayerRef.Bounds(0)(3) - artLayerRef.Bounds(0)(1)
            If w Or h Then
				docRef.ActiveLayer.Invert
			End If
        
            ' The background layer is always non-text, so test to see if it's the background layer.
            ' If it is, then rotate and crop it.
            
            If ( artLayerRef.IsBackgroundLayer ) Then
                docRef.RotateCanvas 45
                docRef.Crop cropBounds
            End If
        End If
    Next
    appRef.Preferences.RulerUnits = strtRulerUnits
Else
    MsgBox "Create a document with a few layers before running this script"
End If

MsgBox "Crop complete"

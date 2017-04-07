' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' This script demonstrates how to stroke and fill the current selection.
' This scripts draws a black stroke around the selection and then
' fills it with red.

Option Explicit

Dim appRef
Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

appRef.DisplayDialogs = 3 ' psDisplayNoDialogs
appRef.Preferences.RulerUnits = 1 ' psPixels

Dim docRef
Dim selRef
Dim strokeColor
Dim fillColor

If ( appRef.Documents.Count > 0 ) Then
    
    Set docRef = appRef.ActiveDocument
    Set selRef = docRef.Selection

    ' Create a new ArtLayer
    docRef.ActiveLayer = docRef.ArtLayers.add

	' Select Area
	Dim docSize, docWidth, docHeight, sPoint, lPoint, selRegion
	
	docWidth = docRef.Width
	docHeight = docRef.Height
    	If docWidth > docHeight Then
		docSize = docHeight
	Else
		docSize = docWidth
	End If

	sPoint = docSize/4
	lPoint = sPoint*3

	selRegion = Array(Array(sPoint, sPoint), _
		Array(lPoint, sPoint), _
		Array(lPoint, lPoint), _
		Array(sPoint, lPoint))
	
	docRef.Selection.Select selRegion

    ' Create the solid color and fill it with a CMYK color
    Set strokeColor = CreateObject( "Photoshop.SolidColor" )
    
    With strokeColor
        .CMYK.Cyan = 0
        .CMYK.Magenta = 0
        .CMYK.Yellow = 0
        .CMYK.Black = 100
    End With

    ' color and width    
    selRef.Stroke strokeColor, 10

    ' Create the solid color and fill it with an RGB color
    Set fillColor = CreateObject( "Photoshop.SolidColor" )

    With fillColor
        .RGB.Red = 255
        .RGB.Green = 0
        .RGB.Blue = 0
    End With

    selRef.Fill fillColor

Else
    MsgBox "Create a document with an active selection before running this script!"
End If

MsgBox "Selection Effects complete"

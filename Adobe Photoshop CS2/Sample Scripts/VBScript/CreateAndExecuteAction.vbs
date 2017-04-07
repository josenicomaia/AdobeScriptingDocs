' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' This script creates an action, which is equivalent to the Mosaic Tiles action
' and executes it.

Option Explicit

Dim appRef
Dim filterDescriptor
Dim retDescriptor
Dim keyTileSizeID
Dim keyGroutWidthID
Dim keyLightenGroutID
Dim eventMosaicID
Dim adesc
Dim actionRef

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

If appRef.Documents.Count <= 0 Then
	Dim fileName
    fileName = appRef.Path & "\Samples\Dune.tif"
    appRef.Open ( fileName )
End If

' create an action and execute it.
keyTileSizeID = appRef.CharIDToTypeID( "TlSz" )
keyGroutWidthID = appRef.CharIDToTypeID( "GrtW" )
keyLightenGroutID = appRef.CharIDToTypeID( "LghG" )
eventMosaicID = appRef.CharIDToTypeID( "MscT" )

Set filterDescriptor = CreateObject( "Photoshop.ActionDescriptor" )
filterDescriptor.PutInteger keyTileSizeID, 12
filterDescriptor.PutInteger keyGroutWidthID, 3
filterDescriptor.PutInteger keyLightenGroutID, 9

Set retDescriptor = appRef.ExecuteAction( eventMosaicID, filterDescriptor, 3 ) ' 3 = psDisplayNoDialogs

MsgBox "Create And Execute Action complete"

' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' This script demonstrates how to open a Photoshop document from the samples folder

Option Explicit

Dim appRef
Dim docRef
Dim fileName

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

fileName = appRef.Path & "\Samples\Dune.tif"

Set docRef = appRef.Open( fileName )

MsgBox "Open Document complete"

' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' This sample script shows how to apply 3 different filters to
' selections in the open document. Adjust the file path for the
' fileName variable as needed to open an appropriate file

Option Explicit

Dim appRef
Dim docRef
Dim textureType
Dim selectionType
Dim fileName
Dim strtRulerUnits

Dim strSamples
Dim strVanishingPoint
Dim strLocString
Dim strArg

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

' We don't want any Photoshop dialogs displayed during automated execution
appRef.DisplayDialogs = 3 ' psDisplayNoDialogs

' We are going to use pixel value inputs, so we need to ensure that
' the current ruler units in Preferences is pixels
strtRulerUnits = appRef.Preferences.RulerUnits
appRef.Preferences.RulerUnits = 1 ' psPixels

strSamples = "$$$/LocalizedFilenames.xml/SourceDirectoryName/id/Extras/[LOCALE]/[LOCALE]_Samples/value=Samples"
strArg = Array(strSamples)
Call getLocString(strSamples)

strVanishingPoint = "$$$/LocalizedFilenames.xml/SourceFileName/id/Extras/[LOCALE]/[LOCALE]_Samples/Vanishing_Point.psd/value=Vanishing Point.psd"
strArg = Array(strVanishingPoint)
Call getLocString(strVanishingPoint)

fileName = appRef.Path & "\" & strSamples & "\" & strVanishingPoint
Set docRef = appRef.Open( fileName )

' Make 3 different selections and apply different filters.
docRef.Selection.Select Array( Array( 0, 485 ), Array( 600, 485 ), Array( 600, 600 ), Array( 0, 600 ) ), 1, 20, True
docRef.ArtLayers( 1 ).ApplyAddNoise 15, 2, False ' 2 = psGaussianNoise

docRef.Selection.Select Array( Array( 120, 20 ), Array( 210, 20 ), Array( 210, 110 ), Array( 120, 110 ) ), 1, 15, False
docRef.ActiveLayer.ApplyDiffuseGlow 9, 12, 15

textureType = 4 ' psTinyLensTexture
docRef.ActiveLayer.ApplyGlassEffect 7, 3, 7, False, textureType

docRef.Selection.Deselect

'Set ruler units back the way we found it
appRef.Preferences.RulerUnits = strtRulerUnits

MsgBox "Filters complete"

' ===============================================
' getLocString functions
' ===============================================
' on localized builds we pull the $$$/Strings from a .dat file, see documentation for more details
Function getLocString(strLocString)

	Dim objWshShell
	Dim myScriptPath
	Dim myFSO
	Dim strJSXFile

	Set objWshShell = WScript.CreateObject("WScript.Shell")
	myScriptPath = objWshShell.CurrentDirectory
	Set myFSO = CreateObject("Scripting.FileSystemObject")
	strJSXFile = myScriptPath + "\localize.jsx"

	strLocString =  appRef.DoJavaScriptFile(strJSXFile,Array(strLocString),1)

End Function

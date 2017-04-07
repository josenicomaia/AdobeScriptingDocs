' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' Apply the "Puzzle (Image)" layer style to the active layer

Option Explicit

Dim appRef
Dim docRef
Dim fileName
Dim artLayerRef

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

Dim strSamples
Dim strVanishingPoint
Dim strPuzzleImage
Dim strLocString
Dim strArg

If appRef.Documents.Count > 0 Then
	Set docRef = appRef.ActiveDocument
Else ' open sample file	
	strSamples = "$$$/LocalizedFilenames.xml/SourceDirectoryName/id/Extras/[LOCALE]/[LOCALE]_Samples/value=Samples"
	strArg = Array(strSamples)
	Call getLocString(strSamples)

	strVanishingPoint = "$$$/LocalizedFilenames.xml/SourceFileName/id/Extras/[LOCALE]/[LOCALE]_Samples/Vanishing_Point.psd/value=Vanishing Point.psd"
	strArg = Array(strVanishingPoint)
	Call getLocString(strVanishingPoint)

	fileName = appRef.Path & "\" & strSamples & "\" & strVanishingPoint
	Set docRef = appRef.Open( fileName )
End If

Set artLayerRef = docRef.ActiveLayer

If artLayerRef.IsBackgroundLayer Then
	artLayerRef.IsBackgroundLayer = False
End If

strPuzzleImage = "$$$/Presets/Styles/DefaultStyles_asl/PuzzleImage=Puzzle (Image)"
strArg = Array(strPuzzleImage)
Call getLocString(strPuzzleImage)

docRef.ActiveLayer.ApplyStyle strPuzzleImage

MsgBox "Apply Style script complete"

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

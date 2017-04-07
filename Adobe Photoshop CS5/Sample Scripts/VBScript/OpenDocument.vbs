' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' This script demonstrates how to open a Photoshop document from the samples folder

Option Explicit

Dim appRef
Dim docRef
Dim fileName

Dim strSamples
Dim strLayerComps
Dim strLocString
Dim strArg

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

strSamples = "$$$/LocalizedFilenames.xml/SourceDirectoryName/id/Extras/[LOCALE]/[LOCALE]_Samples/value=Samples"
strArg = Array(strSamples)
Call getLocString(strSamples)

strLayerComps = "$$$/LocalizedFilenames.xml/SourceFileName/id/Extras/[LOCALE]/[LOCALE]_Samples/Layer_Comps.psd/value=Layer Comps.psd"
strArg = Array(strLayerComps)
Call getLocString(strLayerComps)

fileName = appRef.Path & "\" & strSamples & "\" & strLayerComps
Set docRef = appRef.Open( fileName )

MsgBox "Open Document complete"

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

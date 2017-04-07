' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' This script executes the "Molten Lead" action which is in the
' Photoshop actions palette. If using a non-English version of Photoshop, you
' may have to change the name of the action to match an appropriate action
' in your actions palette.

Option Explicit

Dim appRef
Dim docRef
Dim fileName

Dim strSamples
Dim strLayerComps
Dim strDefaultActions
Dim strMoltenLead
Dim strLocString
Dim strArg

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

If appRef.Documents.Count > 0 Then
    Set docRef = appRef.ActiveDocument
Else ' open sample file

	strSamples = "$$$/LocalizedFilenames.xml/SourceDirectoryName/id/Extras/[LOCALE]/[LOCALE]_Samples/value=Samples"
	strArg = Array(strSamples)
	Call getLocString(strSamples)

	strLayerComps = "$$$/LocalizedFilenames.xml/SourceFileName/id/Extras/[LOCALE]/[LOCALE]_Samples/Layer_Comps.psd/value=Layer Comps.psd"
	strArg = Array(strLayerComps)
	Call getLocString(strLayerComps)

	fileName = appRef.Path & "\" & strSamples & "\" & strLayerComps

	Set docRef = appRef.Open( fileName )
End If

' Execute an action from the action's palette

strDefaultActions = "$$$/Presets/Actions/DefaultActions_atn/DefaultActions=Default Action"
strArg = Array(strDefaultActions)
Call getLocString(strDefaultActions)
strDefaultActions = Replace(strDefaultActions, ".atn","")

strMoltenLead = "$$$/Presets/Actions/SampleActions/MoltenLead=Molten Lead"
strArg = Array(strMoltenLead)
Call getLocString(strMoltenLead)

On Error Resume Next
appRef.DoAction strMoltenLead, strDefaultActions

If Err.Number = 0 Then
	MsgBox "Execute Action complete"
Else
	' MsgBox Err.Description
	MsgBox "Please load " & strDefaultActions & " set and try again."
End If

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

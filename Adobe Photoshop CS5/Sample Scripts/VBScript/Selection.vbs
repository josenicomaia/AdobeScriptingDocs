' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' Make a selection in the active document, copy it to the clipboard and 
' paste it into a new document

Option Explicit

Dim appRef
Dim docRef
Dim fileName
Dim selectionType
Dim selectionRef
Dim newDocRef

Dim strSamples
Dim strVanishingPoint
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

	strVanishingPoint = "$$$/LocalizedFilenames.xml/SourceFileName/id/Extras/[LOCALE]/[LOCALE]_Samples/Vanishing_Point.psd/value=Vanishing Point.psd"
	Call getLocString(strVanishingPoint)

	fileName = appRef.Path & "\" & strSamples & "\" & strVanishingPoint
	Set docRef = appRef.Open( fileName )
End If

appRef.Preferences.RulerUnits = 1 ' psPixels
appRef.DisplayDialogs = 3 ' psDisplayNoDialogs
    
selectionType = 1 ' psReplaceSelection
docRef.Selection.Select Array( Array( 50, 60 ), Array( 150, 60 ), Array( 150, 120 ), Array( 50, 120 ) ), selectionType, 10, False
    
' Get the document selection and copy it to the clipboard.
' If there is a selection the entire document gets copied.
' Then create a new document and paste the selection to the new document.

Set selectionRef = docRef.Selection
docRef.Selection.Copy
    
Set newDocRef = appRef.Documents.Add( 400, 400 )
newDocRef.Paste

MsgBox "Selection complete"

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

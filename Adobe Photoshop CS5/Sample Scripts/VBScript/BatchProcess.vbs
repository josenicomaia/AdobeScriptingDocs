' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' This script will display a directory list dialog, open the folder selected,
' update the document info of the file and save it into a new file and then
' create a thumbnail of the file and save it as a separate jpeg file.

Option Explicit
On Error Resume Next

Dim appRef
Dim docRef
Dim dupDocRef1
Dim dupDocRef2
Dim docInfoRef
Dim fsoRef
Dim folderRef
Dim convertedFolderRef
Dim fileCollection
Dim fileRef
Dim extType
Dim newFileName1
Dim newFileName2
Dim newFolderName
Dim saveOptionsRef
Dim strtRulerUnits
Dim jpgSaveOptions
Dim i

Dim strSamples
Dim strLocString
Dim strArg

Set appRef = CreateObject( "Photoshop.Application" )
appRef.BringToFront
appRef.DisplayDialogs = 3 ' psDisplayNoDialogs
i = 0

Set fsoRef = CreateObject( "Scripting.FileSystemObject" )

strSamples = "$$$/LocalizedFilenames.xml/SourceDirectoryName/id/Extras/[LOCALE]/[LOCALE]_Samples/value=Samples"
strArg = Array(strSamples)
Call getLocString(strSamples)

Set folderRef = fsoRef.GetFolder( appRef.Path & strSamples &"\" )
newFolderName = folderRef & "\VBScriptOut"
saveOptionsRef = 2 ' psDoNotSaveChanges

If fsoRef.FolderExists( newFolderName ) Then
	Set convertedFolderRef = fsoRef.GetFolder( newFolderName )
Else
	Set convertedFolderRef = fsoRef.CreateFolder( newFolderName )
End If

Set fileCollection = folderRef.Files
extType = 2 ' psLowercase

strtRulerUnits = appRef.Preferences.RulerUnits
appRef.Preferences.RulerUnits = 2 ' psInches

For Each fileRef In fileCollection

	' open the file and make first duplicate.
	Set docRef = appRef.Open( fileRef.Path )

	If Err.Number <> 0 Then
		WScript.echo "Unable to open " & fileRef.path
		Err.Clear
	Else
		Set dupDocRef1 = docRef.Duplicate

		' Update the document info of the file.
		Set docInfoRef = dupDocRef1.Info
		docInfoRef.Copyrighted = 1 ' psCopyrightedWork
		docInfoRef.CopyrightNotice = "Copyright 2002, Cool Photoshop Stuff"

		' // This is odd I can't set this. The docs say it is read only.
		If Not docRef.BitsPerChannel = 8 Then
			Call ChangeDocumentDepth( 8 )
			' Set docRef.BitsPerChannel = 8
		End If

		' Create the JPEG options and set the options.
		Set jpgSaveOptions = CreateObject( "Photoshop.JPEGSaveOptions" )
		jpgSaveOptions.EmbedColorProfile = True
		jpgSaveOptions.FormatOptions = 1 ' psStandardBaseline
		jpgSaveOptions.Matte = 1 ' psNoMatte
		jpgSaveOptions.Quality = 1

		' Make up a new name for the new file.
		newFileName1 = convertedFolderRef.Path & "\Temp00" & i
		newFileName1 = newFileName1 & ".jpg"

		' Save with new document information and close the file.
		dupDocRef1.SaveAs newFileName1, jpgSaveOptions, True, extType

		dupDocRef1.Close saveOptionsRef

		' Now create a 1x1 inch thumbnail with a second duplication.
		Set dupDocRef2 = docRef.Duplicate

		' This is odd I can't set this. The docs say it is read only.
		If Not dupDocRef2.BitsPerChannel = 8 Then
			Call ChangeDocumentDepth( 8 )
			' Set docRef.BitsPerChannel = 8
		End If

		dupDocRef2.ResizeImage 1, 1

		' Make up a new name for new thumbnail file.
		newFileName2 = convertedFolderRef.Path & "\Thumbnail00" & i
		newFileName2 = newFileName2 + ".jpg"

		' Save with new document info and close the file.
		dupDocRef2.SaveAs newFileName2, jpgSaveOptions, True, extType
		dupDocRef2.Close saveOptionsRef

		docRef.Close saveOptionsRef

		i = i + 1
	End If
Next

appRef.Preferences.RulerUnits = strtRulerUnits
MsgBox i & " files worked on by Batch Process"

' ===============================================
' Helper functions
' ===============================================
Function ChangeDocumentDepth( ByVal depth )

	Dim id7
	Dim desc3
	Dim id8

	id7 = appRef.CharIDToTypeID( "CnvM" )
	Set desc3 = CreateObject( "Photoshop.ActionDescriptor" )
	id8 = appRef.CharIDToTypeID( "Dpth" )
	Call desc3.PutInteger( id8, depth )
	Call appRef.ExecuteAction( id7, desc3, 3 )

End Function

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

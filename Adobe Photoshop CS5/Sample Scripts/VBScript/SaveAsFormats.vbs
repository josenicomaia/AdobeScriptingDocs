' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' Save a copy of the doc as PDF and JPEG and then save as a Photoshop file with options.

Option Explicit

Dim appRef
Dim docRef
Dim numDocs
Dim extType
Dim fileName
Dim pdfSaveOptions
Dim jpgSaveOptions
Dim psSaveOptions
Dim folderRef
Dim fsoRef
Dim newFolderName
Dim convertedFolderRef
Dim outFileName

Dim strSamples
Dim strLayerComps
Dim strLocString
Dim strArg

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

numDocs = appRef.Documents.Count
extType = 3 ' psUppercase

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

Set fsoRef = CreateObject( "Scripting.FileSystemObject" )
' Set folderRef = fsoRef.GetFolder( appRef.Path & "Samples\" )
Set folderRef = fsoRef.GetFolder( appRef.Path & strSamples & "\" )
newFolderName = folderRef & "\VBScriptOut"

If fsoRef.FolderExists( newFolderName ) Then
    Set convertedFolderRef = fsoRef.GetFolder( newFolderName )
Else
    Set convertedFolderRef = fsoRef.CreateFolder( newFolderName )
End If

' save the doc as PDF
Set pdfSaveOptions = CreateObject( "Photoshop.PDFSaveOptions" )
pdfSaveOptions.AlphaChannels = True
pdfSaveOptions.Annotations = True
pdfSaveOptions.EmbedColorProfile = True
pdfSaveOptions.EmbedFonts = True
pdfSaveOptions.Encoding = 2 ' psPDFJPEG
pdfSaveOptions.Interpolation = False
pdfSaveOptions.JPEGQuality = 7
pdfSaveOptions.Layers = True
pdfSaveOptions.SpotColors = True
pdfSaveOptions.Transparency = False
pdfSaveOptions.UseOutlines = False
pdfSaveOptions.VectorData = True
outFileName = convertedFolderRef.Path & "\Temp.pdf"
docRef.SaveAs outFileName, pdfSaveOptions, True, extType

' now save as JPEG
Set jpgSaveOptions = CreateObject( "Photoshop.JPEGSaveOptions" )
jpgSaveOptions.EmbedColorProfile = True
jpgSaveOptions.FormatOptions = 1 ' psStandardBaseline
jpgSaveOptions.Matte = 1 ' psNoMatte
jpgSaveOptions.Quality = 1
outFileName = convertedFolderRef.Path & "\Temp.jpg"
docRef.SaveAs outFileName, jpgSaveOptions, True, extType

' now save as photoshop with extra options.
Set psSaveOptions = CreateObject( "Photoshop.PhotoshopSaveOptions" )
psSaveOptions.AlphaChannels = True
psSaveOptions.Annotations = True
psSaveOptions.Layers = True
psSaveOptions.SpotColors = True
outFileName = convertedFolderRef.Path & "\Temp.psd"
docRef.SaveAs outFileName, psSaveOptions, False, extType

MsgBox "Save As Formats complete"

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

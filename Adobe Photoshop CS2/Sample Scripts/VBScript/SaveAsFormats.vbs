' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
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

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

numDocs = appRef.Documents.Count
extType = 3 ' psUppercase

If appRef.Documents.Count > 0 Then
    Set docRef = appRef.ActiveDocument
Else ' open sample file
    fileName = appRef.Path & "\Samples\Dune.tif"
    Set docRef = appRef.Open( fileName )
End If

Set fsoRef = CreateObject( "Scripting.FileSystemObject" )
Set folderRef = fsoRef.GetFolder( appRef.Path & "Samples\" )
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

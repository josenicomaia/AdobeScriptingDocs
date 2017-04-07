' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' Store a reference to the document with the name "My Document"
' If the document does not exist, then display an error message.

Option Explicit

Dim appRef
Dim docName
Dim docRef

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

docName = "Untitled-1"

On Error Resume Next

Set docRef = appRef.Documents( docName )

If ( Err.Description = "No such element" ) Then
	MsgBox "Couldn't locate document " & "'" & docName & "'"
Else
	MsgBox "Document Found!"
End If

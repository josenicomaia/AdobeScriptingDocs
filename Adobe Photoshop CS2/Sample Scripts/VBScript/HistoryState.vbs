' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' This script will demonstrate how one can use the HistoryState object to
' "undo" commands.  In this example, the HistoryState is stored, and then
' 2 methods are called simply to modify the HistoryState.  When the ActiveHistoryState
' is reset to the previously saved HistoryState, the actions are "rolled back" to that
' state.  This essentially does an "undo" of the 2 methods called.

Option Explicit

Dim appRef
Dim docRef
Dim currentHistory
Dim fileName

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

If appRef.Documents.Count > 0 Then
    Set docRef = appRef.ActiveDocument
Else ' open sample file
    fileName = appRef.Path & "\Samples\Dune.tif"
    Set docRef = appRef.Open( fileName )
End If
    
Set currentHistory = docRef.HistoryStates( 1 )

' Do a couple of things to change the history state.  This adds
' 2 items to the HistoryStates collection.
docRef.ActiveLayer.AdjustBrightnessContrast 30, 40
docRef.ActiveLayer.AdjustLevels 20, 100, 2, 30, 120

' This rolls back the history to the beginning.
' The "beginning" refers to HistoryStates( 1 )
docRef.ActiveHistoryState = currentHistory

MsgBox "History State complete"

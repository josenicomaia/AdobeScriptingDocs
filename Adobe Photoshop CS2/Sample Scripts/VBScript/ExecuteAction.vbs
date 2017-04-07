' Copyright 2002-2005.  Adobe Systems, Incorporated.  All rights reserved.
' This script executes the "Molten Lead" action which is in the
' Photoshop actions palette. If using a non-English version of Photoshop, you
' may have to change the name of the action to match an appropriate action
' in your actions palette.

Option Explicit

Dim appRef
Dim fileName

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

' Execute an action from the action's palette
appRef.DoAction "Automation workspaces", "Default Actions"

MsgBox "Execute Action complete"

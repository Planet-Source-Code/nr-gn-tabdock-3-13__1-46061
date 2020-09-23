VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITDockMoveEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ===================================================================
' Class: ITDockMoveEvents.cls
' Created: Gary Noble
' Description: Global DockForm Implements
' Uses: This Class Must Be Implemented In Every Docked Form.
'       If It Is Not Then The Form Will not Be Displayed!
'====================================================================
' History:
'   Gary Noble - 5-6-2003
'       Originally This Class Was Implemented To Make It Easier For The
'       User To Handle The Form Sizing Events.
'       A Nicer More Professional Interface
'   Gary Noble - 6-6-2003
'       Add The DockChange Event
' ===================================================================

Public Function Move(Left As Integer, Top As Integer, Bottom As Integer, Right As Integer)
'
End Function

Public Function DockChange(tDockAlign As AlignConstants, tDocked As Boolean)
'
End Function



VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form Layout"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4890
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      Picture         =   "Form3.frx":058A
      ScaleHeight     =   2295
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   300
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ITDockMoveEvents

Private Function ITDockMoveEvents_DockChange(tDockAlign As AlignConstants, tDocked As Boolean) As Variant
    
    Select Case tDockAlign
        Case Is = vbAlignLeft
            Picture1.BackColor = vbWhite
        Case Is = vbAlignRight
            Picture1.BackColor = vbWhite
        Case Is = vbAlignBottom
            Picture1.BackColor = vbRed
        Case Is = vbAlignTop
            Picture1.BackColor = vbYellow
    End Select
       
    If Not tDocked Then Picture1.BackColor = vbYellow
    
    

End Function

Private Function ITDockMoveEvents_Move(Left As Integer, Top As Integer, Bottom As Integer, Right As Integer)
    Picture1.Move Left, Top, Right, Bottom
End Function

'Private Sub Form_Resize()
'
'    On Error Resume Next
'        Picture1.Move 40, Picture1.Top, Me.ScaleWidth - 80, Me.ScaleHeight - (Picture1.Top + 20)
'        '  Picture1.Move 250 + (ScaleWidth / 2) - (Picture1.Width / 2), 200 + (ScaleHeight / 2) - (Picture1.Height / 2)
'    On Error GoTo 0
'
'End Sub

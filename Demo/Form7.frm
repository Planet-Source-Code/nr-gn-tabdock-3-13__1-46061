VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   Caption         =   "Project Group"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "AutocollapseTop toggle"
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Form6 to Left"
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Form2 to Top"
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1155
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   765
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":058A
            Key             =   "closed"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   4260
      _Version        =   393217
      Indentation     =   265
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Implements ITDockMoveEvents
Private Sub Command1_Click()

    ' move form2 to top

    MDIForm1.TabDock.DockChange "Form2", tdAlignTop

End Sub

Private Sub Command2_Click()

    'move form6 to left

    MDIForm1.TabDock.DockChange "Form6", tdAlignLeft

End Sub

Private Sub Command3_Click()
    If MDIForm1.TabDock.AutoCollapseTop Then
        MDIForm1.TabDock.AutoCollapseTop = False
    Else
        MDIForm1.TabDock.AutoCollapseTop = True
    End If
End Sub

Private Sub Form_Load()

    With TreeView1.Nodes
        .Add , , , "Item 1", "closed", "closed"
        .Add 1, tvwChild, , "SubItem 1", "closed", "closed"
        .Add , , , "Item 2", "closed", "closed"
        .Add 3, tvwChild, , "SubItem 1", "closed", "closed"
        .Add 3, tvwChild, , "SubItem 2", "closed", "closed"
        .Add 3, tvwChild, , "SubItem 3", "closed", "closed"
    End With 'TREEVIEW1.NODES

End Sub



Private Function ITDockMoveEvents_DockChange(tDockAlign As AlignConstants, tDocked As Boolean) As Variant
'
End Function

'-- end code

Private Function ITDockMoveEvents_Move(Left As Integer, Top As Integer, Bottom As Integer, Right As Integer)
    TreeView1.Move Left, Top, Right, Bottom
End Function


'Private Sub Form_Resize()
'
'    On Error Resume Next
'    TreeView1.Move MDIForm1.TabDock.DockedFormCaptionOffsetLeft(Me.Name), MDIForm1.TabDock.DockedFormCaptionOffsetTop(Me.Name), Me.ScaleWidth - MDIForm1.TabDock.DockedFormCaptionOffsetRight(Me.Name), Me.ScaleHeight - MDIForm1.TabDock.DockedFormCaptionOffsetBottom(Me.Name)
'    On Error GoTo 0
'
'End Sub

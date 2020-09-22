VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmt 
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text Taker"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "frmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub KeepOnTop(F As Form)
Const SWP_NOMOVE = 2                                        ' Sets the given form On TopMost
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Command1_Click()
CommonDialog1.Filter = "Text files(*.txt)|*.txt|All FIles(*.*)|*.*"
CommonDialog1.ShowOpen
Dim st11 As String * 10000
Open CommonDialog1.FileName For Random As #1 Len = Len(st11)
Get #1, , st11
frmt.Text1.Text = st11
Close #1
End Sub

Private Sub Command2_Click()
CommonDialog1.Filter = "Text files(*.txt)|*.txt|All FIles(*.*)|*.*"
CommonDialog1.ShowSave
Dim s1 As Variant
s1 = frmt.Text1.Text
Open CommonDialog1.FileName For Output As #1
Print #1, s1
Close #1


End Sub

Private Sub Form_Load()
    KeepOnTop frmt

End Sub

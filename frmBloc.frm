VERSION 5.00
Begin VB.Form Block 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Block Website"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "Block"
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
           Open App.Path & "\block.lst" For Append As #1
        Print #1, Text1.Text
                Close #1

 List1.AddItem Text1.Text

End Sub

Private Sub Form_Load()
    KeepOnTop Block
    

Text1.Text = Form1.Text1.Text


    Open App.Path & "\block.lst" For Input As #1
    Do While Not EOF(1)
        Line Input #1, a$
        List1.AddItem a$
    Loop
    Close #1

End Sub

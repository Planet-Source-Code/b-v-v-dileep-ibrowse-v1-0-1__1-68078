VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Favourites"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "Add"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "Favourites"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "URL"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim c As Integer

Sub KeepOnTop(F As Form)
Const SWP_NOMOVE = 2                                        ' Sets the given form On TopMost
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Command1_Click()
           Open App.Path & "\favs.dat" For Append As #1
        Print #1, Text1.Text
                Close #1

 List1.AddItem Text1.Text
 
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    KeepOnTop Form2
c = Form1.tx.Text
Text1.Text = Form1.Text1.Text


    Open App.Path & "\favs.dat" For Input As #1
    Do While Not EOF(1)
        Line Input #1, a$
        List1.AddItem a$
    Loop
    Close #1

End Sub

Private Sub List1_DblClick()

  Text1.Text = List1.List(List1.ListIndex)
        Form2.Visible = False ' Unpresses the button in the tool bar and hides the form
    Dim d As Integer
    d = Form1.TabStrip1.Tabs.Count
    Dim S As Integer
    S = Form1.TabStrip1.Tabs.Count
    Dim s1 As Integer
    s1 = Form1.TabStrip1.SelectedItem.index
    If Form1.wb(c).LocationURL = "about:blank" Then
    Form1.wb(c).Navigate2 (Text1.Text)


Else

    Form1.wb(c).Navigate2 (Text1.Text)
    

End If



    
 

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
           Open App.Path & "\favs.dat" For Append As #1
        Print #1, Text1.Text
                Close #1

End If


End Sub

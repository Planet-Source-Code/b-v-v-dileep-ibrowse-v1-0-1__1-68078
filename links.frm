VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form links 
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Links"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5730
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   5160
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "Save As"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "The links present for selected web page:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "links"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer
Dim colDocuments As Collection

        Dim HTMLlinks As HTMLAnchorElement

Sub KeepOnTop(F As Form)
Const SWP_NOMOVE = 2                                        ' Sets the given form On TopMost
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Command1_Click()
cd.Filter = "Text File(*.txt)|*.txt"
On Error Resume Next
cd.ShowSave
Open cd.FileName For Output As #2
          GetFrames
           Print #2, "Links extracted from" & " " & Form1.Text1.Text & ":"
           
           For i = 1 To colDocuments.Count

    For Each HTMLlinks In colDocuments.Item(i).links

            Print #2, HTMLlinks.href
            
            Next HTMLlinks
Next i
Close #2

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    KeepOnTop links
c = Form1.tx.Text
Dim HTMLdoc As HTMLDocument
            Dim STRtxt As String
    ' List the links.
    Dim ln As Integer
    ln = Form1.tx.Text
    
          GetFrames
           For i = 1 To colDocuments.Count

    For Each HTMLlinks In colDocuments.Item(i).links

            List1.AddItem HTMLlinks.href
            
            Next HTMLlinks
Next i

End Sub
Private Sub GetFrames()
'Purpose: searches all frames in the document(s) and adds them to colDocuments

Dim i As Integer, j As Integer
j = 1
On Error Resume Next
Set colDocuments = New Collection
colDocuments.Add Form1.wb(Form1.TabStrip1.Tabs(c).Tag).Document

Do While j < colDocuments.Count + 1
    For i = 0 To colDocuments.Item(j).frames.Length - 1
        colDocuments.Add colDocuments.Item(j).frames.Item(i).Document
    Next i
    j = j + 1
Loop
End Sub

Private Sub List1_DblClick()
  Text1.Text = List1.List(List1.ListIndex)

    Dim d As Integer
    d = Form1.TabStrip1.Tabs.Count
    Dim S As Integer
    S = Form1.TabStrip1.Tabs.Count
    Dim s1 As Integer
    s1 = Form1.TabStrip1.SelectedItem.index
    If Form1.wb(s1).LocationURL = "about:blank" Then
    Form1.wb(d).Navigate2 (Text1.Text)

Form1.TabStrip1.Tabs(d).Caption = Text1.Text

Else

    Form1.wb(S).Navigate2 (Text1.Text)
End If
Unload Me

End Sub

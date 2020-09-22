VERSION 5.00
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   3240
   ClientLeft      =   -45
   ClientTop       =   -330
   ClientWidth     =   4605
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin vbskfr2.Skinner Skinner1 
      Left            =   4080
      Top             =   2400
      _ExtentX        =   1270
      _ExtentY        =   1270
      SysDisableSkinCaption=   "&Disable Skin"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Enter Key"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblReg 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Get Key"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   2
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID :"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Code"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
Unload Me
Form1.Show

End Sub

Private Sub Command2_Click()
  TrialTime Me, "The trial of " & Me.Caption & " has expired. Please register this product to get the full version.", "Trial Expired", vbCritical, 50, True
  'Activates the trial counter. True to count up and False to reset the Trial count
   Label9.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
Form1.Show
Unload Me
End Sub


Private Sub Command7_Click()
 Dim CharLength As Integer
    Dim i As Integer
    Dim Char As String
    Dim Key As Variant
    
    CharLength% = Len(Text1.Text)
    If CharLength = 0 Then Exit Sub
    For i% = 1 To CharLength%
        Char$ = Mid(Text1.Text, i, 1)
         Key = Key & (Asc(Char) Xor CharLength%)
    Next i%
    If Key = Text2.Text Then
        MsgBox ("Thank You For Registering"), , "Thank You"
        TrialTime Me, "", "", "", 0, False
                  
    Else
        MsgBox ("The key you have entered seems to incorrect"), , "Incorrect"
        Me.Height = 5040
    End If
End Sub


Private Sub Form_Load()
Label2.Caption = "Get Key"

 On Error Resume Next
 Label9.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
 If Label9.Caption = 15 Then
 Command2.Enabled = False
 Label8.Caption = "TRIAL TIME EXPIRED"
 End If
 Label8.Caption = "Used " & Label9.Caption & " Times out of 15"
p.Value = 15 - Label9.Caption

'Display trial count
   'Make forms width 1860
    If GetSetting(Me.Name, "Trial", "TimesOpen") = "." Then
    lblReg.FontSize = 8: lblReg.Caption = "Already Registered!"
    Command7.Visible = False
Command1.Caption = "OK"
        Label2.Visible = False
'this means they registered the program.
    'put the full version code here
    'Example: me.unload  and frmFullVersion.show
    Else
        lblReg.FontSize = 8: lblReg.Caption = "Unregistered!"
    Command7.Visible = True
Command1.Caption = "Cancel"
Label2.Visible = True
    End If
End Sub
Private Function TrialTime(TheForm As Form, TrialOverMSG As String, TrialOverMSGTitle As String, TrialOverMSGType As String, TrialCount As Integer, Work As Boolean)

    If Not Work Then SaveSetting TheForm.Name, "Trial", "TimesOpen", ".": End
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting TheForm.Name, "Trial", "TimesOpen", Val(GetSetting(TheForm.Name, "Trial", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened

    If GetSetting(TheForm.Name, "Trial", "TimesOpen") > TrialCount Then SaveSetting TheForm.Name, "Trial", "TimesOpen", TrialCount: MsgBox TrialOverMSG, TrialOverMSGType, TrialOverMSGTitle: End
'If the amount of times open is > then the TrialCount..
'Reset it to the number in TrialCount specified
'Display a message and terminate the program
End Function


Private Sub Label2_Click()
Dim shellsuccess11 As Long

shellsuccess11 = ShellExecute(fH, "Open", "mailto:bvvdileep@Gmail.com?subject=Registration key for IBrowse.", 0&, 0&, 10)

End Sub



VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   8880
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   7560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer2 
      Interval        =   5
      Left            =   240
      Top             =   960
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7815
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10001
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tx 
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   6240
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
      Height          =   255
      Left            =   10920
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9120
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form14.frx":113A
      Left            =   7080
      List            =   "Form14.frx":1150
      TabIndex        =   6
      Text            =   "Select Search engine"
      Top             =   960
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   4695
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   5655
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   11775
      ExtentX         =   20770
      ExtentY         =   9975
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   661
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      TabMinWidth     =   1413
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":1184
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":1726
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":1E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":27BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":2F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":3862
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":4020
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":4962
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":5134
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":5B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":6358
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":6C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":750C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":7DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":86C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":8F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":9874
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":A14E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":AE20
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":BA4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":C328
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":CF7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":D3CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":D81E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":DC70
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":E0C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":E600
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   2340
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   4128
      ButtonWidth     =   1535
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "Stop"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "History"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Favourites"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mail"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Menu fl 
      Caption         =   "File"
      Begin VB.Menu nt 
         Caption         =   "New Tab"
         Shortcut        =   ^T
      End
      Begin VB.Menu nw 
         Caption         =   "New Window"
         Shortcut        =   ^W
      End
      Begin VB.Menu op 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu sa 
         Caption         =   "Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu ct 
         Caption         =   "Close Tab"
      End
      Begin VB.Menu clal 
         Caption         =   "Close All"
      End
      Begin VB.Menu sp11 
         Caption         =   "-"
      End
      Begin VB.Menu ps 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu pr 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu sp12 
         Caption         =   "-"
      End
      Begin VB.Menu ed12 
         Caption         =   "Edit"
      End
      Begin VB.Menu ap 
         Caption         =   "Advaned Properties"
      End
      Begin VB.Menu ew 
         Caption         =   "Email Web page"
      End
      Begin VB.Menu pr11 
         Caption         =   "Page Properties"
      End
      Begin VB.Menu ex 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu ed 
      Caption         =   "Edit"
      Begin VB.Menu cp 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu sa11 
         Caption         =   "Select All"
      End
      Begin VB.Menu sp14 
         Caption         =   "-"
      End
      Begin VB.Menu fn 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu vw 
      Caption         =   "View"
      Begin VB.Menu zi 
         Caption         =   "Zoom In"
      End
      Begin VB.Menu ts 
         Caption         =   "Text Size"
         Begin VB.Menu sm1111 
            Caption         =   "Smallest"
         End
         Begin VB.Menu sm11 
            Caption         =   "Small"
         End
         Begin VB.Menu me 
            Caption         =   "Medium"
         End
         Begin VB.Menu la 
            Caption         =   "Large"
         End
         Begin VB.Menu lr 
            Caption         =   "Largest"
         End
      End
      Begin VB.Menu hs 
         Caption         =   "History"
      End
      Begin VB.Menu fs 
         Caption         =   "Full Screen"
      End
   End
   Begin VB.Menu bl 
      Caption         =   "Block"
      Begin VB.Menu im 
         Caption         =   "Images"
      End
      Begin VB.Menu pb 
         Caption         =   "Popup Blocker"
      End
      Begin VB.Menu bs 
         Caption         =   "Block Script Errors"
      End
   End
   Begin VB.Menu fv 
      Caption         =   "Favourites"
      Begin VB.Menu af 
         Caption         =   "Add to Favourites"
         Shortcut        =   ^A
      End
      Begin VB.Menu vf 
         Caption         =   "Organize Favourites"
      End
   End
   Begin VB.Menu tl 
      Caption         =   "Tools"
      Begin VB.Menu sc 
         Caption         =   "Shortcuts"
         Begin VB.Menu dt 
            Caption         =   "Desktop"
         End
         Begin VB.Menu md 
            Caption         =   "My Documents"
         End
         Begin VB.Menu mp 
            Caption         =   "My Pictures"
         End
         Begin VB.Menu wn 
            Caption         =   "WINDOWS"
         End
         Begin VB.Menu fn11 
            Caption         =   "Fonts"
         End
         Begin VB.Menu sm111111 
            Caption         =   "Start Menu Folder"
         End
      End
      Begin VB.Menu sp1212 
         Caption         =   "-"
      End
      Begin VB.Menu de 
         Caption         =   "Delete"
         Begin VB.Menu ck11 
            Caption         =   "Cookies"
         End
      End
      Begin VB.Menu te 
         Caption         =   "Text Taker"
      End
      Begin VB.Menu ln 
         Caption         =   "Links Graber"
      End
      Begin VB.Menu sm 
         Caption         =   "Send Mail"
      End
      Begin VB.Menu sp2222 
         Caption         =   "-"
      End
      Begin VB.Menu af11 
         Caption         =   "Auto Refresh"
         Begin VB.Menu ds11 
            Caption         =   "Disabled"
         End
         Begin VB.Menu sp12121212 
            Caption         =   "-"
         End
         Begin VB.Menu e10 
            Caption         =   "Every 10 Sec."
         End
         Begin VB.Menu e20 
            Caption         =   "Every 20 Sec"
         End
         Begin VB.Menu em 
            Caption         =   "Every Minute"
         End
      End
   End
   Begin VB.Menu tr 
      Caption         =   "Translate"
      Begin VB.Menu eg 
         Caption         =   "English To German"
      End
      Begin VB.Menu ec 
         Caption         =   "English To Chinese"
      End
      Begin VB.Menu ef 
         Caption         =   "English To Frensh"
      End
      Begin VB.Menu ep 
         Caption         =   "English To Portugese"
      End
      Begin VB.Menu ei 
         Caption         =   "English To Italian"
      End
      Begin VB.Menu ej 
         Caption         =   "English To Japanese"
      End
      Begin VB.Menu er 
         Caption         =   "English To Russian"
      End
      Begin VB.Menu es 
         Caption         =   "English To Spanish"
      End
   End
   Begin VB.Menu op11 
      Caption         =   "Options"
      Begin VB.Menu io 
         Caption         =   "Internet Options"
      End
   End
   Begin VB.Menu hl 
      Caption         =   "Help"
      Begin VB.Menu cn 
         Caption         =   "Contents"
      End
      Begin VB.Menu rm 
         Caption         =   "Read Me"
      End
      Begin VB.Menu rs 
         Caption         =   "Register"
      End
      Begin VB.Menu ab 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Needs a reference to Microsoft HTML Object Library from Project->References


Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameW" (ByVal lpBuffer As Long, ByRef nSize As Long) As Long
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Const CSIDL_COOKIES As Long = &H21
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim blim As Boolean
Dim i As Integer
Dim hl11 As New Htmlhelp
Dim blim1 As Boolean
Dim strRestrict As Variant
Dim iCount As Integer
Private FirstTime As Boolean
Dim HTMLdoc As HTMLDocument
Dim r As Boolean
Dim p As Boolean
Private colDocuments As Collection
Dim CurrentBrowser As Integer
Dim numtabs As Integer
Dim scrt As Boolean
Private m_iCount As Long
Dim ini As String
Dim find As Boolean
Dim iniFile As String
Dim wb11 As Boolean
Private Sub ac_Click()
ac.Checked = Not ac.Checked
blim1 = im.Checked

End Sub

Private Sub ab11_Click()
Block.Show

End Sub

Private Sub ab_Click()
frmAbout.Show

End Sub

Private Sub af_Click()
Form2.Show

End Sub

Private Sub ap_Click()
frmAdv.Visible = True
frmAdv.Caption = Form1.Text1.Text

End Sub

Private Sub bs_Click()
bs.Checked = Not bs.Checked
scrt = bs.Checked

End Sub

Private Sub bw_Click()
wb11 = im.Checked

bw.Checked = Not bw.Checked
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Refresh

End Sub

Private Sub ck_Click()
wb(TabStrip1.SelectedItem.index).Navigate getSpecialFolder(&H21)
End Sub

Private Sub clal_Click()
Dim i As Integer
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Stop
For i = 1 To TabStrip1.Tabs.Count
If wb.Count = 1 Then
    wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate "about:blank"
   Else
   wb(TabStrip1.Tabs(CurrentBrowser).Tag).Visible = False
    Unload wb(TabStrip1.Tabs(CurrentBrowser).Tag)
    TabStrip1.Tabs.Remove (CurrentBrowser)
    i = 0
    numtabs = numtabs - 1
    If TabStrip1.Tabs.Count < CurrentBrowser Then
        CurrentBrowser = CurrentBrowser - 1
    End If
    SelectTab (CurrentBrowser)
End If
Next

End Sub

Private Sub cn_Click()
With hl11
.CHMFile = App.Path & "\Help.chm"
.HHWindow = ""
.HHDisplayContents
End With

End Sub

Private Sub Command1_Click()
On Error Resume Next
wb(TabStrip1.SelectedItem.index).Navigate2 Text1.Text

End Sub
Private Sub ck11_Click()
Dim sPath As String
sPath = Space(260)
Call SHGetSpecialFolderPath(0, sPath, CSIDL_COOKIES, False)
sPath = Left$(sPath, InStr(sPath, vbNullChar) - 1) & "\*.txt*"
On Error Resume Next
Kill sPath
End Sub

Private Sub Command2_Click()
If Combo1.ListIndex = 0 Then
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate2 ("http://www.google.com/search?q=" & Text2.Text)
End If
If Combo1.ListIndex = 1 Then
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate2 ("http://search.yahoo.com/search?p=" & Text2.Text)

End If
If Combo1.ListIndex = 2 Then
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate2 ("http://search.lycos.com/default.asp?query=" & Text2.Text)

End If
If Combo1.ListIndex = 3 Then
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate2 ("http://web.ask.com/web?q=" & Text2.Text)
End If

If Combo1.ListIndex = 4 Then
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate2 ("http://www.altavista.com/web/results?q=" & Text2.Text)

End If

If Combo1.ListIndex = 5 Then
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate2 ("http://www.answers.com/main/ntquery?s=" & Text2.Text)

End If



End Sub

Private Sub cp_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT


End Sub

    

Private Sub ct_Click()
Dim i As Integer
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Stop
If wb.Count = 1 Then
    wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate "about:blank"
   Else
   wb(TabStrip1.Tabs(CurrentBrowser).Tag).Visible = False
    Unload wb(TabStrip1.Tabs(CurrentBrowser).Tag)
    TabStrip1.Tabs.Remove (CurrentBrowser)
    i = 0
    numtabs = numtabs - 1
    If TabStrip1.Tabs.Count < CurrentBrowser Then
        CurrentBrowser = CurrentBrowser - 1
    End If
    SelectTab (CurrentBrowser)
End If




End Sub
Public Sub SelectTab(index As Integer)
If index > numtabs Then
    Call MsgBox("The tab that you selected" & vbCrLf & "is outof range" _
                , vbCritical, "Error Selecting Tab")
    Exit Sub
End If
TabStrip1.Tabs(index).Selected = True
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Visible = False
wb(TabStrip1.Tabs(index).Tag).Visible = True
wb(TabStrip1.Tabs(index).Tag).ZOrder
wb(TabStrip1.Tabs(index).Tag).SetFocus
CurrentBrowser = index
CurrentAddress = wb(TabStrip1.Tabs(CurrentBrowser).Tag).LocationURL
End Sub

Private Sub ds11_Click()
ds11.Checked = Not ds11.Checked
e10.Checked = False
e20.Checked = False
em.Checked = False
Timer1.Enabled = False

End Sub

Private Sub dt_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate getSpecialFolder(&H10)
End Sub

Private Sub e10_Click()
e10.Checked = Not e10.Checked
e20.Checked = False
em.Checked = False
ds11.Checked = False
Timer1.Interval = 1000

End Sub

Private Sub e20_Click()
e20.Checked = Not e20.Checked
e10.Checked = False
em.Checked = False
ds11.Checked = False
Timer1.Interval = 200

End Sub

Private Sub ec_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate "http://babelfish.altavista.com/babelfish/urltrurl?&lp=en_zh&tt=url&url=" & Text1.Text

End Sub

Private Sub ed12_Click()
Dim cap As String
On Error GoTo esource
frmSource.Text1.Text = wb(TabStrip1.Tabs(CurrentBrowser).Tag).Document.documentElement.innerHTML
frmSource.Caption = cap
frmSource.WebBrowser1.Navigate wb(TabStrip1.Tabs(CurrentBrowser).Tag).LocationURL

frmSource.Show
esource:

End Sub

Private Sub ef_Click()
wb(TabStrip1.SelectedItem.index).Navigate "http://babelfish.altavista.com/babelfish/urltrurl?&lp=en_fr&tt=url&url=" & Text1.Text

End Sub

Private Sub eg_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate "http://babelfish.altavista.com/babelfish/urltrurl?&lp=en_de&tt=url&url=" & Text1.Text


End Sub

Private Sub ei_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate "http://babelfish.altavista.com/babelfish/urltrurl?&lp=en_it&tt=url&url=" & Text1.Text

End Sub

Private Sub ej_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate "http://babelfish.altavista.com/babelfish/urltrurl?&lp=en_ja&tt=url&url=" & Text1.Text

End Sub

Private Sub em_Click()
em.Checked = Not em.Checked
e10.Checked = False
e20.Checked = False
ds11.Checked = False
Timer1.Interval = 1000

End Sub

Private Sub ep_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate "http://babelfish.altavista.com/babelfish/urltrurl?&lp=en_pt&tt=url&url=" & Text1.Text

End Sub

Private Sub er_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate "http://babelfish.altavista.com/babelfish/urltrurl?&lp=en_ru&tt=url&url=" & Text1.Text

End Sub

Private Sub es_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate "http://babelfish.altavista.com/babelfish/urltrurl?&lp=en_es&tt=url&url=" & Text1.Text

End Sub

Private Sub ew_Click()
'Dim shellsuccess11 As Long

'shellsuccess11 = ShellExecute(fH, "Open", "mailto:", 0&, 0&, 10)
    ShellExecute 0&, vbNullString, "mailto:?subject= &body=" & Text1.Text, vbNullString, vbNullString, vbHide


End Sub
Public Function SendMail(ByVal MailAddress As String, ByVal MailSubject As String, ByVal MailBody As String, ByVal MailAttach As String)

End Function
Private Sub fn_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).SetFocus
    SendKeys "^f"

End Sub

Private Sub fn11_Click()
wb(TabStrip1.SelectedItem.index).Navigate getSpecialFolder(&H14)
End Sub



Private Sub Form_Load()

Dim i12 As String
TabStrip1.Tabs.Item(1).Tag = 1
CurrentBrowser = 1
r = True
i = 0
wb(1).Navigate "about:blank"
ct.Enabled = False

    End Sub



Private Sub fs_Click()
frmFull.Show
frmFull.Caption = Form1.Text1.Text

End Sub

Private Sub fv11_Click()
wb(TabStrip1.SelectedItem.index).Navigate getSpecialFolder(&H6)
End Sub

Private Sub hs_Click()
 hs.Checked = Not hs.Checked
 If hs.Checked = True Then
 Hist.Visible = True

 Toolbar1.Buttons(8).Value = tbrPressed
 Else
 Toolbar1.Buttons(8).Value = tbrUnpressed
 Hist.Visible = False
 
End If
End Sub

Private Sub hs11_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate getSpecialFolder(&H22)
End Sub

Private Sub im_Click()
im.Checked = Not im.Checked
blim = im.Checked
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Refresh

End Sub

Private Sub io_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)

End Sub

Private Sub lr_Click()
On Error Resume Next
wb(TabStrip1.Tabs(CurrentBrowser).Tag).ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull

End Sub

Private Sub mc_Click()
wb(TabStrip1.SelectedItem.index).Navigate getSpecialFolder(&H1)
End Sub

Private Sub md_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate getSpecialFolder(&H5)
End Sub

Private Sub mp_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate getSpecialFolder(&H27)
End Sub

Private Sub ot_Click()
frmOption.Show

End Sub

Private Sub pb_Click()
pb.Checked = Not pb.Checked
If pb.Checked = True Then
p = True
Else
p = False
End If
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Refresh

End Sub

Private Sub rm_Click()
ShellExecute hWnd, "open", App.Path & "\Ibrowse.rtf", vbNullString, vbNullString, conSwNormal

End Sub

Private Sub rs_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate "http://www.ibrowse.tk"

End Sub

Private Sub sm11_Click()
On Error Resume Next
wb(TabStrip1.Tabs(CurrentBrowser).Tag).ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull

End Sub

Private Sub sm1111_Click()
On Error Resume Next
wb(TabStrip1.Tabs(CurrentBrowser).Tag).ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull

End Sub

Private Sub sm111111_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate getSpecialFolder(&HB)
End Sub

Private Sub te11_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate2 "http://translate.google.com/translate_t"

End Sub

Private Sub TabStrip1_GotFocus()
Dim i As Integer
For i = 1 To TabStrip1.Tabs.Count
    If TabStrip1.Tabs(i).Selected = True Then
        SelectedTab = i
    Else
        wb(TabStrip1.Tabs(i).Tag).Visible = False
    End If
Next i
With wb(TabStrip1.Tabs(SelectedTab).Tag)
.Visible = True
.ZOrder
.SetFocus
CurrentAddress = .LocationURL
End With
CurrentBrowser = SelectedTab
Form1.Caption = TabStrip1.Tabs(CurrentBrowser).Caption
Text1.Text = wb(TabStrip1.Tabs(CurrentBrowser).Tag).LocationURL
tx.Text = CurrentBrowser
If CurrentBrowser = 1 Then
ct.Enabled = False
Else
ct.Enabled = True
End If

End Sub

Private Sub Text2_LostFocus()
If Text2.Text = "" Then
Text2.Text = "Enter search query"
End If


End Sub

Private Sub la_Click()
On Error Resume Next
wb(TabStrip1.Tabs(CurrentBrowser).Tag).ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull

End Sub

Private Sub ln_Click()
links.Show

End Sub

Private Sub me_Click()
On Error Resume Next
wb(TabStrip1.Tabs(CurrentBrowser).Tag).ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull

End Sub

Private Sub sm_Click()
Dim shellsuccess As Long

shellsuccess = ShellExecute(fH, "Open", "mailto:", 0&, 0&, 10)

End Sub

Private Sub te_Click()
frmt.Show

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate2 Text1.Text
    End If
End Sub

Private Sub nt_Click()
'TabStrip1.Tabs.Add
'TabStrip1.Tabs(TabStrip1.Tabs.Count).Tag = TabStrip1.Tabs.Count

'TabStrip1.Tabs(TabStrip1.Tabs.Count).Caption = "Untitled" & TabStrip1.Tabs.Count
'Load wb(TabStrip1.Tabs.Count)


 '   wb(TabStrip1.Tabs.Count).Navigate2 ("about:blank")
NewTab ("about:blank")
tx.Text = CurrentBrowser

End Sub

Private Sub nw_Click()
On Error Resume Next
Dim frm As Form1
Set frm = New Form1
Set ppDisp = frm.wb(TabStrip1.Tabs(CurrentBrowser).Tag).object

frm.Show

End Sub

Private Sub op_Click()
On Error Resume Next
Set HTMLdoc = Form1.wb(TabStrip1.Tabs(CurrentBrowser).Tag).Document
CommonDialog1.Filter = "All Internet Files (*.htm,*.html,*.asp,*.shtml,*.js,*.dhtml) | *.htm;*.html;*.asp;*.shtml;*.js;*.dhtml|" & _
"All Supported Picture formats(*.gif,*.tif,*.pcd,*.jpg,*.wmf,*.tga,*.jpeg,*.ras,*.png,*.eps,*.bmp,*.pcx)|*.gif;*.tif;*.pcd;*.jpg;*.wmf;*.tga;*.jpeg;*.ras;*.png;*.eps;*.bmp;*.pcx|" & _
        "Text formats (*.txt,*.doc)|*.txt;*.doc|" & _
        "All files (*.*)|*.*|"
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate CommonDialog1.FileName
End Sub


Private Sub pr_Click()
wb(TabStrip1.SelectedItem.index).ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub pr11_Click()
On Error Resume Next


wb(TabStrip1.SelectedItem.index).ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub ps_Click()
wb(TabStrip1.SelectedItem.index).ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub sa_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub sa11_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub Timer1_Timer()
wb(TabStrip1.SelectedItem.index).Refresh
End Sub

Private Sub Timer2_Timer()
If scrt = True Then
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Silent = True

Timer2.Enabled = False
Else
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Silent = True
End If
End Sub

Private Sub tm_Click()
wb(TabStrip1.SelectedItem.index).Navigate getSpecialFolder(&H20)
End Sub

Private Sub tw_Click()
wb(TabStrip1.SelectedItem.index).Navigate2 "http://translate.google.com/translate_t"

End Sub

Private Sub vf_Click()
Form2.Show

End Sub

Private Function isThere(file1 As String)
On Error GoTo F
Open file1 For Input As #1
find = True
Close #1
Exit Function
F:
find = False
End Function

Private Sub wb_BeforeNavigate2(index As Integer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Text1.Text = wb(TabStrip1.Tabs(CurrentBrowser).Tag).LocationURL

End Sub

Private Sub wb_CommandStateChange(index As Integer, ByVal Command As Long, ByVal Enable As Boolean)
Select Case Command
Case 1 'Forward
Toolbar1.Buttons.Item("Forward").Enabled = Enable
Case 2 'Back
Toolbar1.Buttons.Item("Back").Enabled = Enable
End Select
End Sub

Private Sub wb_NavigateComplete2(index As Integer, ByVal pDisp As Object, URL As Variant)
Text1.Text = wb(TabStrip1.Tabs(CurrentBrowser).Tag).LocationURL

   
    

End Sub

Private Sub wb_OnToolBar(index As Integer, ByVal ToolBar As Boolean)
    
    On Error Resume Next
If p = True Then
StatusBar1.Panels.Item(3).Text = "Popups Blocked"
If ToolBar = False Then
        Unload Me
    End If
    Else
    StatusBar1.Panels.Item(3).Text = "Popups Unblocked"
End If

End Sub

Private Sub wb_TitleChange(index As Integer, ByVal Text As String)
wb(index).Tag = Text
Me.Caption = wb(index).Tag
End Sub
Private Sub wb_StatusTextChange(index As Integer, ByVal Text As String)
StatusBar1.Panels(1).Text = Text
End Sub

Private Sub wb_ProgressChange(index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
If Progress = -1 Then ProgressBar1.Value = 100 'the name of the progress bar is "ProgressBar1".        Label1.Caption = "Done"
ProgressBar1.Visible = False
'This makes the progress bar disappear after the page is loaded.

If Progress > 0 And ProgressMax > 0 Then
ProgressBar1.Visible = True

ProgressBar1.Value = Progress * 100 / ProgressMax

End If
Exit Sub
End Sub
Private Sub wb_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
If p = True Then
On Error Resume Next
Dim frm As Form1
Cancel = IsPopupWindow
If Cancel = False Then
Set frm = New Form1
Set ppDisp = frm.WebBrowser1.objectfrm.Show
End If
Else
End If

End Sub
End Sub
Private Function IsPopupWindow() As Boolean
On Error Resume Next
If WebBrowser1.Document.activeElement.tagName = "BODY" Or WebBrowser1.Document.activeElement Then
IsPopupWindow = True
Else
IsPopupWindow = False
End If
End Function
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Caption = "Back" Then
On Error Resume Next

wb(TabStrip1.Tabs(CurrentBrowser).Tag).GoBack
End If
If Button.Caption = "Forward" Then
On Error Resume Next

wb(TabStrip1.Tabs(CurrentBrowser).Tag).GoForward
End If
If Button.Caption = "Refresh" Then
On Error Resume Next

wb(TabStrip1.Tabs(CurrentBrowser).Tag).Refresh
End If

If Button.Caption = "Stop" Then
On Error Resume Next

wb(TabStrip1.Tabs(CurrentBrowser).Tag).Stop
End If

If Button.Caption = "Home" Then
On Error Resume Next

wb(TabStrip1.Tabs(CurrentBrowser).Tag).GoHome
End If
If Button.Caption = "Search" Then
wb(TabStrip1.Tabs(CurrentBrowser).Tag).GoSearch
End If


If Button.Caption = "Favourites" Then
On Error Resume Next
Dim shellHelper As New ShellUIHelper
    Dim strLocationName, strLocationURL As String

    strLocationName = Form1.wb(TabStrip1.Tabs(CurrentBrowser).Tag).LocationName
    strLocationURL = Form1.wb(TabStrip1.Tabs(CurrentBrowser).Tag).LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName

End If

If Button.Caption = "Print" Then
wb(TabStrip1.Tabs(CurrentBrowser).Tag).ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End If
If Button.Caption = "History" Then
  hs.Checked = Not hs.Checked
 If hs.Checked = True Then
 Toolbar1.Buttons(8).Value = tbrPressed
Hist.Visible = True
 
 Else
 Toolbar1.Buttons(10).Value = tbrUnpressed

Hist.Visible = False

End If
End If
If Button.Caption = "Mail" Then
Dim shellsuccess As Long

shellsuccess = ShellExecute(fH, "Open", "mailto:", 0&, 0&, 10)
End If
If Button.Caption = "Edit" Then
Dim cap As String
On Error GoTo esource
frmSource.Text1.Text = wb(TabStrip1.Tabs(CurrentBrowser).Tag).Document.documentElement.innerHTML
frmSource.Caption = cap
frmSource.Show
esource:
End If


End Sub

Private Sub GetFrames()
'Purpose: searches all frames in the document(s) and adds them to colDocuments
On Error Resume Next
Dim i11 As Integer, j As Integer
j = 1

Set colDocuments = New Collection
colDocuments.Add wb(TabStrip1.Tabs(CurrentBrowser).Tag).Document

Do While j < colDocuments.Count + 1
    For i11 = 0 To colDocuments.Item(j).frames.Length - 1
        colDocuments.Add colDocuments.Item(j).frames.Item(i11).Document
    Next i11
    j = j + 1
Loop
End Sub

Private Sub wb_DownloadComplete(index As Integer)
On Error Resume Next
 Dim HTMLdoc1 As HTMLDocument
 Dim hl As HTMLImg
 If blim = True Then
 Set HTMLdoc1 = wb(TabStrip1.Tabs(CurrentBrowser).Tag).Document
GetFrames
For i = 1 To colDocuments.Count
    For Each hl In colDocuments.Item(i).images
    hl.src = ""
       
Next hl
Next i

StatusBar1.Panels.Item(2).Text = "Images Blocked"

Else
StatusBar1.Panels.Item(2).Text = "Images Unblocked"

End If
TabStrip1.SelectedItem.Caption = Form1.Caption
If p = True Then
StatusBar1.Panels.Item(3).Text = "Popups Blocked"
    Else
    StatusBar1.Panels.Item(3).Text = "Popups Unblocked"
End If

End Sub

Private Sub wn_Click()
wb(TabStrip1.Tabs(CurrentBrowser).Tag).Navigate getSpecialFolder(&H24)
End Sub

Private Sub zi_Click()
Dim i1 As Integer

wb(TabStrip1.Tabs(CurrentBrowser).Tag).ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DODEFAULT






End Sub
Private Function FirstIndex() As Integer
'Finds first unused or released index of the browser control array
Dim found As Boolean
Dim i As Integer
For Each object In wb()
    If object.index <> i Then
        found = True
        Exit For
    End If
    i = i + 1
Next
If found = True Then
    FirstIndex = i
Else
FirstIndex = numtabs
End If
End Function

Public Sub NewTab(URL As String)
'Adds a new tab and browser to the control
Dim tmp As Integer
numtabs = TabStrip1.Tabs.Count
    TabStrip1.Tabs.Add
    numtabs = numtabs + 1
    CurrentBrowser = numtabs
    With TabStrip1.Tabs(CurrentBrowser)
    .Caption = "New Page"
    .Selected = False
    .Tag = FirstIndex
    tmp = .Tag
End With
' load a new browser control with the first available index
Load wb(TabStrip1.Tabs(CurrentBrowser).Tag)
wb(tmp).Visible = True
wb(tmp).ZOrder

SelectTab CurrentBrowser
wb(tmp).Navigate URL
' you can add specific options here to tailor this code to your needs
' for example:
'
' select case Options
'       case 1          'Navigate to homepage
'           browser(tmp).Navigate "www.mypage.com"
'       case 2          'Navigate to GOOGLE
'           browser(tmp).Navigate "www.google.com"
' end select

End Sub

Public Sub DeleteTab()
End Sub

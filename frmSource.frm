VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmSource 
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8550
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   6015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   8175
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   150
      Left            =   2760
      TabIndex        =   0
      Top             =   2400
      Width           =   135
      ExtentX         =   238
      ExtentY         =   265
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
      Location        =   ""
   End
   Begin VB.Menu fl 
      Caption         =   "File"
      Begin VB.Menu sa 
         Caption         =   "Save As"
      End
      Begin VB.Menu sp11 
         Caption         =   "-"
      End
      Begin VB.Menu ps 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu pt 
         Caption         =   "Print"
      End
      Begin VB.Menu sp22 
         Caption         =   "-"
      End
      Begin VB.Menu et 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSource"
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
Unload Me
End Sub



Private Sub fn_Click()
CodeMax1.ExecuteCmd cmCmdFind

End Sub

Private Sub Form_Load()
KeepOnTop frmSource
WebBrowser1.Silent = True

End Sub

Private Sub sa_Click()
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT

End Sub


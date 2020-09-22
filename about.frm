VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Private Sub Form_Load()
lblTitle.Caption = "IBrowse V1.0" & vbCrLf & "Author:Dileep" & vbCrLf & "Credits:" & "Sherif Rofael(for Win API tutorial)" & vbCrLf & "Ryan Sever(for Multi tabbed browsing concept)" & vbCrLf & "Icons by Xp Icandy" & vbCrLf


End Sub
        
 
 

Private Sub Form_Click()
Unload Me

End Sub


Private Sub Label1_Click()
Dim shellsuccess As Long

shellsuccess = ShellExecute(fH, "Open", "http//:www.IBrowse.tk", 0&, 0&, 10)

End Sub

Private Sub lblTitle_Click()
Unload Me

End Sub

Private Sub Picture1_Click()
Unload Me

End Sub

Private Sub Timer1_Timer()
    If lblTitle.Top < Picture1.Height - Picture1.Height - lblTitle.Height Then
        lblTitle.Top = Picture1.Height - 1
        lblTitle.Top = lblTitle.Top - 5
    Else
        lblTitle.Top = lblTitle.Top - 10
    End If

End Sub

Private Sub TabStrip1_Click()

End Sub

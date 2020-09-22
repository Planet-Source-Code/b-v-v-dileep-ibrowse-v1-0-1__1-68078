VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdv 
   Caption         =   "Advanced Properties"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4500
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdv.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdv.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdv.frx":11B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7011
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmAdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ln As Integer
Dim colDocuments As Collection
Dim c As Integer
Private Sub Form_Load()
   c = Form1.tx.Text
  tv.Nodes.Clear
 On Error Resume Next
Dim HTMLdoc As HTMLDocument
        Dim HTMLlinks As HTMLAnchorElement
        Dim HTMLlinks1 As HTMLImg
           

Dim HTMLlinks2 As HTMLImg
    ' List the links.
    ln = Form1.tx.Text
    Set HTMLdoc = Form1.wb(ln).Document
  On Error Resume Next
        
            Dim nodX As Node
   Set nodX = tv.Nodes.Add(, tvwParent, "R4", "Title")
    Set nodX = tv.Nodes.Add("R4", tvwChild, HTMLdoc.Title, HTMLdoc.Title)
    

   Set nodX = tv.Nodes.Add(, tvwParent, "R", "Links")
GetFrames
                For i = 1 To colDocuments.Count

    For Each HTMLlinks In colDocuments.Item(i).links

            Set nodX = tv.Nodes.Add("R", tvwChild, HTMLlinks.href, HTMLlinks.href, 1)
            
            
            Next HTMLlinks
Next i

      Set nodX = tv.Nodes.Add(, tvwParent, "R1", "Images")
GetFrames
           For i = 1 To colDocuments.Count

    For Each HTMLlinks1 In colDocuments.Item(i).images

           Set nodX = tv.Nodes.Add("R1", tvwChild, HTMLlinks1.href, HTMLlinks1.href, 2)
           
            
            Next HTMLlinks1
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

Sub AddToTreeView(mText As String, mParent As String, Optional mImage As Integer)
    
    On Error GoTo ErrPoint
    Dim tvNode As Node
    If mImage = 0 Then
        Set tvNode = tv.Nodes.Add(mParent, tvwChild, Right(mText, 20), mText)
    Else
        Set tvNode = tv.Nodes.Add(mParent, tvwChild, Right(mText, 20), mText, mImage)
    End If

ErrPoint:

End Sub

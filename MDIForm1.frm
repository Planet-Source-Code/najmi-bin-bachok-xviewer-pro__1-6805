VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "XViewer Pro?"
   ClientHeight    =   4140
   ClientLeft      =   810
   ClientTop       =   1005
   ClientWidth     =   4860
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuFileExit_Click()
 End
End Sub

Private Sub mnuFileNew_Click()
 
 Dim frm1 As Object
 
 Set frm1 = New Form1
 Load frm1
 frm1.Show
 
End Sub

Private Sub mnuFileOpen_Click()
 
On Error GoTo ed

Dim cmddlg32 As New cCommonDialog
Dim fs() As String
Dim indeks As Long
Dim frm1 As Object
With cmddlg32
 
 .hwnd = Me.hwnd
 .FileTitle = ""
 .Filename = ""
 .CancelError = False
 .Filter = "All Graphics Files|*.bmp;*.cur;*.dib;*.emf;*.gif;*.icl;*.ico;*.jfif;*.jpe;*.jpeg;*.jpg;*.wmf"
 .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_FILEMUSTEXIST _
       Or OFN_HIDEREADONLY
 .ShowOpen
 
 If .Filename = "" Then Exit Sub
 
 .ParseMultiFileName .Filename, fs, indeks

 If indeks = 1 Then
     Set frm1 = New Form1
     Load frm1
     frm1.Show
     frm1.Picture1.Picture = LoadPicture(fs(1), 0, 0, 0, 0)
     frm1.Picture1.AutoSize = True
 End If
 
 If indeks > 1 Then
 For i = 1 To indeks
    If i + 1 > indeks Then Exit For

     Set frm1 = New Form1
     Load frm1
     frm1.Show
     frm1.Picture1.Picture = LoadPicture(fs(i + 1), 0, 0, 0, 0)
     frm1.Picture1.AutoSize = True
     
     DoEvents
     
  With frm1
 
  .VScroll1.Max = .Picture1.Height
  .VScroll1.Value = 0
  
  .HScroll1.Max = .Picture1.Width
  .HScroll1.Value = 0
  
  End With
  
 Next
 
 End If
 
End With
  
  Exit Sub
  
ed:
  MsgBox "Error while opening file.", vbCritical, Err.Description
  Exit Sub
  
End Sub

Private Sub mnuHelpAbout_Click()

Dim msg As String
msg = " Brought for free by ©MasterX Artwork"
msg = msg & vbCrLf & " Version " & App.Major & "." & App.Minor & vbCrLf
msg = msg & " E-mail: gonejoe@hotmail.com" & vbCrLf
msg = msg & " Web: http://www.geocities.com/m_bachok" & vbCrLf
msg = msg & " Happy Viewing :)"

MsgBox msg, vbInformation, "About Xviewer Pro?"

End Sub

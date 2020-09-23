VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "XViewer"
   ClientHeight    =   4440
   ClientLeft      =   1050
   ClientTop       =   1125
   ClientWidth     =   5370
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   5370
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.VScrollBar VScroll1 
         Height          =   2895
         LargeChange     =   50
         Left            =   4440
         SmallChange     =   20
         TabIndex        =   2
         Top             =   960
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   100
         Left            =   600
         SmallChange     =   20
         TabIndex        =   3
         Top             =   3840
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "X"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3840
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   240
         Negotiate       =   -1  'True
         ScaleHeight     =   169
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   217
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
   End
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
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditClear 
         Caption         =   "C&lear"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditClearClipBoard 
         Caption         =   "Clear Clipboard"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowArrangeIcon 
         Caption         =   "&Arrange Icon"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileVert 
         Caption         =   "&Tile Vertical"
      End
      Begin VB.Menu mnuWindowTileHorz 
         Caption         =   "Tile &Horizontal"
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Stupid XViewer © MasterX Artwork
'Web site: http://www.geocities.com/m_bachok
'E-mail: gonejoe@hotmail.com
'********************************************************************************
'Just a simple picture box placed the form
'With a little enhancements made, this little tweaking viewer works
'A little different then fully functional graphics editor, this viewer
'allows you to view Bitmap based graphics, Metafiles and Icons.
'There are some Paste, Cut Clear, and Copy function added to this project
'List of Image Type supported by Picture box
' 1) vbPicTypeIcon = Windows Native Icon
' 2) vbPicTypeBitmap = Bitmap graphics e.g JPEG, GIF, Windows Bitmap
' 3) vbPictypeMetafile = Vector Metafile
' 4) vbPictypeEMetafile = Enhanced Vector Metafile



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 Select Case KeyCode
  
  Case vbKeyLeft:
   Me.HScroll1.SetFocus
    
  Case vbKeyRight:
    Me.HScroll1.SetFocus
   
  Case vbKeyUp:
    Me.VScroll1.SetFocus
  
  Case vbKeyDown:
   Me.VScroll1.SetFocus
   
  Case vbKeyPageDown:
   Me.VScroll1.SetFocus
  
  Case vbKeyPageUp:
   Me.VScroll1.SetFocus
   
 End Select
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

 Select Case KeyCode
  
  Case vbKeyLeft:
   Me.HScroll1.SetFocus
    
  Case vbKeyRight:
    Me.HScroll1.SetFocus
   
  Case vbKeyUp:
    Me.VScroll1.SetFocus
  
  Case vbKeyDown:
   Me.VScroll1.SetFocus
   
  Case vbKeyPageDown:
   Me.VScroll1.SetFocus
  
  Case vbKeyPageUp:
   Me.VScroll1.SetFocus
    
 End Select
 
End Sub

Private Sub Form_Load()
 
 With Me
 
  .Frame1.Move 0, 0
  .Frame1.Width = Me.ScaleWidth - 100
  .Frame1.Height = Me.ScaleHeight - 100
  
  .Picture1.Move 0, 0
  .Picture1.AutoSize = True
  .Picture1.ScaleMode = vbPixels
  
  .VScroll1.Max = .Picture1.Height
  .VScroll1.Value = 0
  
  .HScroll1.Max = .Picture1.Width
  .HScroll1.Value = 0
  
  .VScroll1.Move .Frame1.Width - .VScroll1.Width, .Frame1.Top, _
                 .VScroll1.Width, .Frame1.Height - .Command1.Height
                 
  .HScroll1.Move .Frame1.Left, .Frame1.Height - .HScroll1.Height, _
                 .Frame1.Width - .Command1.Width, .HScroll1.Height
                 
  .Command1.Move .Frame1.Width - .Command1.Width, .Frame1.Height - .Command1.Height
  
 End With

End Sub

Private Sub Form_Resize()

 On Error Resume Next
 
 With Me
 
  .Frame1.Move 0, 0
  .Frame1.Width = Me.ScaleWidth - 100
  .Frame1.Height = Me.ScaleHeight - 100
  
  .Picture1.AutoSize = True
  .Picture1.Refresh
  .Picture1.ScaleMode = vbPixels
  
  .VScroll1.Max = .Picture1.Height
  .VScroll1.LargeChange = .VScroll1.Max / 2
  .VScroll1.Value = 0
  
  .HScroll1.Max = .Picture1.Width
  .HScroll1.LargeChange = .HScroll1.Max / 2
  .HScroll1.Value = 0
  
  .VScroll1.Move .Frame1.Width - .VScroll1.Width, .Frame1.Top, _
                 .VScroll1.Width, .Frame1.Height - .Command1.Height
                 
  .HScroll1.Move .Frame1.Left, .Frame1.Height - .HScroll1.Height, _
                 .Frame1.Width - .Command1.Width, .HScroll1.Height
                 
  .Command1.Move .Frame1.Width - .Command1.Width, .Frame1.Height - .Command1.Height
  
 End With

End Sub

Private Sub Frame1_Click()
 
 Me.Picture1.SetFocus
 
End Sub

Private Sub HScroll1_Change()
 
 On Error Resume Next
 
 With Me

   .Picture1.Left = -.HScroll1.Value
   .Picture1.AutoSize = True
   
 End With
 
End Sub

Private Sub HScroll1_GotFocus()

 Me.SetFocus
 
End Sub

Private Sub HScroll1_KeyDown(KeyCode As Integer, Shift As Integer)

 Select Case KeyCode
  
  Case vbKeyLeft:
  Me.HScroll1.SetFocus
    
  Case vbKeyRight:
    Me.HScroll1.SetFocus
   
  Case vbKeyUp:
    Me.VScroll1.SetFocus
  
  Case vbKeyDown:
   Me.VScroll1.SetFocus
 
  Case vbKeyPageDown:
   Me.VScroll1.SetFocus
  
  Case vbKeyPageUp:
   Me.VScroll1.SetFocus
      
 End Select
 
End Sub




Private Sub mnuEdit_Click()
  On Error Resume Next
  If Me.Picture1.Picture.Handle = 0 Then
   
   Me.mnuEditCopy.Enabled = False
   Me.mnuEditCut.Enabled = False
   Me.mnuEditClear.Enabled = False
   
  Else
   
   Me.mnuEditCopy.Enabled = True
   Me.mnuEditCut.Enabled = True
   Me.mnuEditClear.Enabled = True
   
  End If

  If Clipboard.GetData Then
    
    Me.mnuEditPaste.Enabled = True
  
  Else
  
    Me.mnuEditPaste.Enabled = False
  
  End If
   
End Sub

Private Sub mnuEditClear_Click()

 If Me.Picture1.Picture.Handle = 0 Then Exit Sub
 
 Me.Picture1.Picture = Nothing
 Me.VScroll1.Value = 0
 Me.HScroll1.Value = 0
 Me.Picture1.AutoSize = False
 
End Sub

Private Sub mnuEditClearClipBoard_Click()

 On Error Resume Next
 
 Clipboard.Clear
 
End Sub

Private Sub mnuEditCopy_Click()

 If Me.Picture1.Picture.Handle = 0 Then Exit Sub

 If Me.Picture1.Picture.Type = vbPicTypeIcon Then
  MsgBox "Cannot copy Icon", vbCritical, Err.Description
  Exit Sub
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeBitmap Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFBitmap
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeMetafile Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFMetafile
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeEMetafile Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFEMetafile
 End If

End Sub

Private Sub mnuEditCut_Click()
 
 If Me.Picture1.Picture.Type = vbPicTypeIcon Then
  MsgBox "Cannot copy Icon. Save as bitmap first.", vbCritical, Err.Description
  Exit Sub
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeBitmap Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFBitmap
   Me.Picture1.Picture = Nothing
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeMetafile Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFMetafile
   Me.Picture1.Picture = Nothing
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeEMetafile Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFEMetafile
   Me.Picture1.Picture = Nothing
 End If

 
End Sub

Private Sub mnuEditPaste_Click()

On Error Resume Next

If Clipboard.GetData Then
 With Me
 'Load the image file from clipboard
 Me.Picture1.Picture = Clipboard.GetData
 Me.Picture1.AutoSize = True

   Me.Picture1.Move 0, 0
  
   Me.VScroll1.Max = .Picture1.Height
   Me.VScroll1.SmallChange = .VScroll1.Max / 50
   Me.VScroll1.LargeChange = .VScroll1.Max / 20

   Me.VScroll1.Value = 0
  
   Me.HScroll1.Max = .Picture1.Width
   Me.HScroll1.SmallChange = .HScroll1.Max / 50
   Me.HScroll1.LargeChange = .HScroll1.Max / 20
   Me.HScroll1.Value = 0
   
 End With
End If

End Sub

Private Sub mnuFile_Click()

  If Me.Picture1.Picture.Handle = 0 Then
   
   Me.mnuFileSave.Enabled = False
   
  Else
   
   Me.mnuFileSave.Enabled = True
   
  End If

End Sub

Private Sub mnuFileClose_Click()
 Unload Me
End Sub

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
  Me.Picture1.Picture = LoadPicture(.Filename, 0, 0, 0, 0)
  Exit Sub
 End If
 
 For i = 1 To indeks
 
    If i + 1 > indeks Then Exit Sub
    
    Me.Picture1.Picture = LoadPicture(fs(2))
    
    If i + 1 > 2 Then
    
     Set frm1 = New Form1
     Load frm1
     frm1.Show
     frm1.Picture1.Picture = LoadPicture(fs(i + 1))
     frm1.Picture1.AutoSize = True
     
     DoEvents
     
    End If
 Next
 
End With
Exit Sub

ed:
  MsgBox "Error while opening file.", vbCritical, Err.Description
  Exit Sub
  
End Sub

Private Sub mnuFileSave_Click()

 On Error GoTo ed
 
 Dim cmddlgSave As New cCommonDialog
 
 If Me.Picture1.Picture.Handle = 0 Then Exit Sub
 
 With cmddlgSave

 If Me.Picture1.Picture.Type = vbPicTypeIcon Then
 
  .CancelError = True
  .DialogTitle = "Save Icon"
  .Filename = "Untitled"
  .DefaultExt = "ico"
  .Filter = "Icon|*.ico"
  .hwnd = 0
  .flags = 4 Or OFN_OVERWRITEPROMPT
  .InitDir = ""
  .ShowSave
  If .Filename <> "" Then SavePicture Me.Picture1.Picture, .Filename
  Exit Sub
   
 End If
  
  If Me.Picture1.Picture.Type = vbPicTypeMetafile Then
 
  .CancelError = True
  .DialogTitle = "Save MetaFile"
  .Filename = "Untitled"
  .DefaultExt = "wmf"
  .Filter = "Windows Metafile|*.wmf"
  .hwnd = 0
  .flags = 4 Or OFN_OVERWRITEPROMPT
  .InitDir = ""
  .ShowSave
  If .Filename <> "" Then SavePicture Me.Picture1.Picture, .Filename
  Exit Sub
   
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeEMetafile Then
 
  .CancelError = True
  .DialogTitle = "Save Enhanced MetaFile"
  .Filename = "Untitled"
  .DefaultExt = "emf"
  .Filter = "Enhanced Metafile|*.emf"
  .hwnd = 0
  .flags = 4 Or OFN_OVERWRITEPROMPT
  .InitDir = ""
  .ShowSave
  If .Filename <> "" Then SavePicture Me.Picture1.Picture, .Filename
  Exit Sub
   
 End If

 If Me.Picture1.Picture.Type = vbPicTypeBitmap Then
 
  .CancelError = True
  .DialogTitle = "Save Bitmap Graphics"
  .Filename = "Untitled"
  .DefaultExt = "bmp"
  .Filter = "Bitmap|*.bmp"
  .hwnd = 0
  .flags = 4 Or OFN_OVERWRITEPROMPT
  .InitDir = ""
  .ShowSave
  If .Filename <> "" Then SavePicture Me.Picture1.Picture, .Filename
  Exit Sub
  
 End If
 
 End With

ed:
 If Err.Number = 32755 Then Exit Sub
 MsgBox "Error while saving file", vbCritical, Err.Description
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

Private Sub mnuWindowArrangeIcon_Click()

 MDIForm1.Arrange vbArrangeIcons
 
End Sub

Private Sub mnuWindowCascade_Click()

 MDIForm1.Arrange vbCascade
 
End Sub

Private Sub mnuWindowTileHorz_Click()
 
 MDIForm1.Arrange vbTileHorizontal

End Sub

Private Sub mnuWindowTileVert_Click()

 MDIForm1.Arrange vbTileVertical

End Sub

Private Sub Picture1_Change()
  
 On Error Resume Next
 
 Dim picType As String
 
 With Me
 
  .Frame1.Move 0, 0
  .Frame1.Width = Me.ScaleWidth - 100
  .Frame1.Height = Me.ScaleHeight - 100
  
  .Picture1.AutoSize = True
  .Picture1.Refresh
  .Picture1.ScaleMode = vbPixels
  
  .VScroll1.Max = .Picture1.Height
  .VScroll1.LargeChange = .VScroll1.Max / 2
  .VScroll1.Value = 0
  
  .HScroll1.Max = .Picture1.Width
  .HScroll1.LargeChange = .HScroll1.Max / 2
  .HScroll1.Value = 0
  
  .VScroll1.Move .Frame1.Width - .VScroll1.Width, .Frame1.Top, _
                 .VScroll1.Width, .Frame1.Height - .Command1.Height
                 
  .HScroll1.Move .Frame1.Left, .Frame1.Height - .HScroll1.Height, _
                 .Frame1.Width - .Command1.Width, .HScroll1.Height
                 
  .Command1.Move .Frame1.Width - .Command1.Width, .Frame1.Height - .Command1.Height
  
 End With
 
 If Me.Picture1.Picture.Handle = 0 Then
   Me.Caption = "No Picture Loaded"
   Exit Sub
 End If
 
     Select Case Me.Picture1.Picture.Type
     Case vbPicTypeIcon: picType = "(Windows Icon)"
     Case vbPicTypeBitmap: picType = "(Bitmap)"
     Case vbPicTypeEMetafile: picType = "(Enhanced Metafile)"
     Case vbPicTypeMetafile: picType = "(Metafile)"
     Case vbPicTypeNone: picType = "(Unknown Graphics Type)"
    End Select

   Me.Caption = CLng(Me.Picture1.ScaleX(Picture1.Picture.Width)) & _
                "x" & CLng(Me.Picture1.ScaleY(Picture1.Picture.Height)) & _
                " - " & picType
 
End Sub

Private Sub Picture1_DblClick()

 mnuFileOpen_Click
 
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

 Select Case KeyCode
  
  Case vbKeyLeft:
   Me.HScroll1.SetFocus
    
  Case vbKeyRight:
    Me.HScroll1.SetFocus
   
  Case vbKeyUp:
    Me.VScroll1.SetFocus
  
  Case vbKeyDown:
   Me.VScroll1.SetFocus
   
  Case vbKeyPageDown:
   Me.VScroll1.SetFocus
  
  Case vbKeyPageUp:
   Me.VScroll1.SetFocus
    
 End Select
 
End Sub

Private Sub VScroll1_Change()
 
 On Error Resume Next
 
 With Me

   .Picture1.Top = -.VScroll1.Value
   .Picture1.AutoSize = True

 End With
  
End Sub

Private Sub VScroll1_GotFocus()

 Me.SetFocus
 
End Sub

Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)

 Select Case KeyCode
  
  Case vbKeyLeft:
  Me.HScroll1.SetFocus
    
  Case vbKeyRight:
    Me.HScroll1.SetFocus
   
  Case vbKeyUp:
    Me.VScroll1.SetFocus
  
  Case vbKeyDown:
   Me.VScroll1.SetFocus
   
  Case vbKeyPageDown:
   Me.VScroll1.SetFocus
  
  Case vbKeyPageUp:
   Me.VScroll1.SetFocus
    
   
 End Select
 
End Sub

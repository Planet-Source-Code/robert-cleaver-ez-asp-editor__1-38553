VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ez ASP Editor"
   ClientHeight    =   4860
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmtest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpad 
      Height          =   3180
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   7530
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ez ASP Editor"
      Height          =   3555
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   7755
   End
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   60
      TabIndex        =   2
      Top             =   3540
      Width           =   7755
      Begin VB.Timer Timer1 
         Interval        =   4000
         Left            =   5730
         Top             =   240
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Save As"
         Height          =   435
         Left            =   4530
         TabIndex        =   13
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton Command7 
         Height          =   435
         Left            =   3000
         Picture         =   "frmtest.frx":1042
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Height          =   435
         Left            =   2520
         Picture         =   "frmtest.frx":1258
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Height          =   435
         Left            =   2040
         Picture         =   "frmtest.frx":146E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Height          =   435
         Left            =   1560
         Picture         =   "frmtest.frx":1750
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Height          =   435
         Left            =   1080
         Picture         =   "frmtest.frx":1A32
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Height          =   435
         Left            =   600
         Picture         =   "frmtest.frx":1D44
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   435
         Left            =   120
         Picture         =   "frmtest.frx":1FEE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   60
      TabIndex        =   10
      Top             =   4200
      Width           =   7755
      Begin VB.CommandButton Command8 
         Caption         =   "Exit"
         Height          =   255
         Left            =   6840
         TabIndex        =   12
         Top             =   240
         Width           =   795
      End
      Begin VB.Label statusbar 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6615
      End
   End
   Begin MSComDlg.CommonDialog fileDLG 
      Left            =   7140
      Top             =   3030
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu sbmnuOpen 
         Caption         =   "Open... "
         Shortcut        =   ^O
      End
      Begin VB.Menu sbmnuSaveAs 
         Caption         =   "Save As... "
         Shortcut        =   ^R
      End
      Begin VB.Menu sbmnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu sbmnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu sbmnuMinimize 
         Caption         =   "Minimize"
         Shortcut        =   ^M
      End
      Begin VB.Menu sbmnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu sbmnuClose 
         Caption         =   "Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      NegotiatePosition=   1  'Left
      Begin VB.Menu sbmnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu sbmnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^P
      End
      Begin VB.Menu sbmnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu sbmnuReplace 
         Caption         =   "Replace"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu sbmnuCodeSnippets 
         Caption         =   "Code Snippets"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu sbmnuAbout 
         Caption         =   "About"
         Begin VB.Menu sbmnuAboutAuthor 
            Caption         =   "Author"
            Shortcut        =   ^A
         End
         Begin VB.Menu sbmnuAboutProgram 
            Caption         =   "Program"
            Shortcut        =   ^Q
         End
      End
      Begin VB.Menu sbmnuSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu sbmnuHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Saved = False Then
        frmOpenSave.Show vbModal
    Else
        Call OpenNew
    End If
End Sub

Private Sub Command2_Click()
    Call SaveASP
End Sub

Private Sub Command3_Click()
    Dim Replace1$, Replace2$
    
    Replace1$ = InputBox$("Enter the Word you would Like to Replace:", "Replace Step 1")
    Replace2$ = InputBox$("Enter the Word you Would like to Replace" & Replace1 & " With:", "Replace Step 2")
    Me.txtpad.Text = Replace(Me.txtpad.Text, Replace1$, Replace2$)
End Sub

Private Sub Command4_Click()
    Clipboard.Clear
    Clipboard.SetText txtpad.Text
End Sub

Private Sub Command5_Click()
    Dim PasteWhat$
    PasteWhat$ = Clipboard.GetText
    txtpad.Text = txtpad.Text & PasteWhat$
End Sub

Private Sub Command6_Click()
    frmSnippets.Show vbModal
End Sub

Private Sub Command7_Click()
    If Saved = False Then
        frmYouSure.Show vbModal
    Else
        txtpad.Text = ("")
    End If
End Sub

Private Sub Command8_Click()
    If Saved = False Then
        frmSaveCancel.Show vbModal
    Else
        End
    End If
End Sub

Private Sub Command9_Click()
    Call SaveAs
End Sub

Private Sub Form_Load()
Dim cmd As String, StrBuff As String
Dim tFile As Long

    cmd = Command$
    
    If Len(Trim(cmd)) <= 0 Then
        Saved = True
        SetStatus ("Idle...")
        Exit Sub
    Else
        tFile = FreeFile
        Open cmd For Binary Access Read As #tFile
            StrBuff = Space(LOF(tFile))
            Get #tFile, , StrBuff
        Close #tFile
        OldPath = cmd
        '
        txtpad.Text = StrBuff
        StrBuff = ""
        cmd = ""
        Saved = True
        SetStatus ("By Robert Cleaver")
    End If
If Right$(lzpath, 1) = "\" Then fixpath = lzpath Else fixpath = lzpath & "\"
Dim ans
Dim IconPath As String, ProgPath As String
    ans = MsgBox("This will now install the needed keys for your new file type " _
    & vbNewLine & "Do you want to carry on", vbYesNo Or vbQuestion)
    If ans = vbNo Then: Unload Form1: Exit Sub
    IconPath = fixpath(App.Path) & "appIcon.ico"
    ProgPath = App.Path & "\ASPEDIT.EXE"
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".snp"
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".snp\DefaultIcon"
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".snp\shell"
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".snp\shell\open"
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".snp\shell\open\command"
    Reg32Mod.SaveString HKEY_CLASSES_ROOT, ".snp\DefaultIcon", "", IconPath
    Reg32Mod.SaveString HKEY_CLASSES_ROOT, ".snp\shell\open\command", "", Chr(34) & ProgPath & Chr(34) & " %1"
End Sub

Private Sub Label1_Click()

End Sub

Private Sub sbmnuAboutAuthor_Click()
    frmAboutAuth.Show vbModal
End Sub

Private Sub sbmnuClose_Click()
    If Saved = False Then
        frmSaveCancel.Show vbModal
    Else
        End
    End If
End Sub

Private Sub sbmnuCodeSnippets_Click()
    frmSnippets.Show vbModal
End Sub

Private Sub sbmnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtpad.Text
End Sub

Private Sub sbmnuMinimize_Click()
    Me.WindowState = 1
End Sub

Private Sub sbmnuOpen_Click()
    Call OpenNew
End Sub

Private Sub sbmnuPaste_Click()
    Dim PasteWhat$
    PasteWhat$ = Clipboard.GetText
    txtpad.Text = txtpad.Text & PasteWhat$
End Sub

Private Sub sbmnuReplace_Click()
    Dim Replace1$, Replace2$
    
    Replace1$ = InputBox$("Enter the Word you would Like to Replace:", "Replace Step 1")
    Replace2$ = InputBox$("Enter the Word you Would like to Replace" & Replace1 & " With:", "Replace Step 2")
    Me.txtpad.Text = Replace(Me.txtpad.Text, Replace1$, Replace2$)
End Sub

Private Sub sbmnuSave_Click()
    Call SaveASP
End Sub

Private Sub sbmnuSaveAs_Click()
    Call SaveAs
End Sub

Private Sub Timer1_Timer()
    SetStatus ("Idle...")
End Sub

Private Sub txtpad_Change()
    Saved = False
    SetStatus ("Typing...")
End Sub

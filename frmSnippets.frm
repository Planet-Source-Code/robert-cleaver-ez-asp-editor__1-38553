VERSION 5.00
Begin VB.Form frmSnippets 
   Caption         =   "Code Snippets"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSnippets.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Snippet List"
      Height          =   1905
      Left            =   2550
      TabIndex        =   1
      Top             =   30
      Width           =   1755
      Begin VB.ListBox lstSnippets 
         Height          =   1620
         Left            =   60
         TabIndex        =   8
         Top             =   210
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Snippet Source"
      Height          =   1905
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2595
      Begin VB.TextBox txtSnippet 
         Height          =   1635
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   210
         Width           =   2385
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   30
      TabIndex        =   2
      Top             =   1860
      Width           =   4275
      Begin VB.CommandButton Command3 
         Caption         =   "Load"
         Height          =   345
         Left            =   2670
         TabIndex        =   6
         Top             =   180
         Width           =   1515
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Insert"
         Height          =   345
         Left            =   810
         TabIndex        =   5
         Top             =   180
         Width           =   1605
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Copy"
         Height          =   345
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Height          =   405
         Left            =   2490
         TabIndex        =   3
         Top             =   120
         Width           =   105
      End
   End
   Begin VB.Frame Frame5 
      Height          =   375
      Left            =   30
      TabIndex        =   9
      Top             =   2340
      Width           =   4275
      Begin VB.Label snpStatus 
         Caption         =   "No Snippets Loaded"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmSnippets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Code() As String

Private Sub Command1_Click()
    Clipboard.Clear
    Clipboard.SetText txtSnippet.Text
End Sub

Private Sub Command2_Click()
    frmMain.txtpad.Text = frmMain.txtpad.Text & txtSnippet.Text
End Sub

Private Sub Command3_Click()
    Call LoadSnippets
End Sub

Function LoadSnippets()
Dim Directory$, Buffer$, Buffer2$, SnippetName$
Dim LoopOne%, LoopTwo%, FileCount%, FreeX
Open App.Path & "\snippets\snippets.conf" For Input As #1
    Input #1, FileCount%
Close #1
ReDim Code(FileCount%) As String
Directory$ = App.Path & "\snippets\"
For LoopOne% = 1 To FileCount%
    FreeX = FreeFile
    Open Directory$ & LoopOne% & ".snp" For Input As FreeX
        For LoopTwo% = 1 To 2
            If LoopTwo% = 1 Then
                Input #FreeX, Buffer$
                SnippetName$ = Buffer$
                Buffer2$ = Buffer$
                SnippetName = Replace(SnippetName, "<!-- ", "")
                SnippetName = Replace(SnippetName, " -->", "")
                lstSnippets.AddItem (SnippetName$)
                Buffer$ = ("")
            ElseIf LoopTwo% = 2 Then
                While Not EOF(FreeX)
                    Input #FreeX, Buffer$
                    Buffer2$ = Buffer2$ & vbNewLine & Buffer$
                    Buffer$ = ("")
                Wend
                Code(LoopOne%) = Buffer2$
                Buffer2$ = ("")
            End If
        Next LoopTwo%
    Close FreeX
Next LoopOne%
frmSnippets.snpStatus.Caption = ("Snippets Loaded: " & FileCount%)
End Function

Private Sub lstSnippets_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstSnippets.ListCount = 0 Then Exit Sub
    txtSnippet.Text = Code(lstSnippets.ListIndex + 1)
End Sub

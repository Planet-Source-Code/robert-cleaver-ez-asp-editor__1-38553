VERSION 5.00
Begin VB.Form frmYouSure 
   Caption         =   "Are you Sure?"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYouSure.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Save"
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3855
      Begin VB.Label Label1 
         Caption         =   "You have not Saved Yet. If you Continue All That you have Changed will Be Lost."
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Continue"
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   2580
         TabIndex        =   3
         Top             =   240
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmYouSure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call SaveASP
End Sub

Private Sub Command2_Click()
    frmMain.txtpad.Text = ("")
    Unload Me
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

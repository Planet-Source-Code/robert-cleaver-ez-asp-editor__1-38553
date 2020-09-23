VERSION 5.00
Begin VB.Form frmAboutAuth 
   Caption         =   "About"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "About the Author"
      Height          =   2175
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   405
         Left            =   3480
         TabIndex        =   6
         Top             =   1710
         Width           =   1275
      End
      Begin VB.PictureBox Picture1 
         Height          =   1815
         Left            =   120
         Picture         =   "frmAboutAuth.frx":0000
         ScaleHeight     =   1755
         ScaleWidth      =   1965
         TabIndex        =   1
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label4 
         Caption         =   $"frmAboutAuth.frx":DD86
         Height          =   1155
         Left            =   2220
         TabIndex        =   5
         Top             =   870
         Width           =   2445
      End
      Begin VB.Label Label3 
         Caption         =   "About:"
         Height          =   285
         Left            =   2220
         TabIndex        =   4
         Top             =   660
         Width           =   2085
      End
      Begin VB.Label Label2 
         Caption         =   "Age: 16"
         Height          =   285
         Left            =   2220
         TabIndex        =   3
         Top             =   450
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "Name: Robert Cleaver"
         Height          =   285
         Left            =   2220
         TabIndex        =   2
         Top             =   240
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmAboutAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub


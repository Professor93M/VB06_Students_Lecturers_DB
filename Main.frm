VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·’›Õ… «·—∆Ì”Ì…"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "«·«”« –…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "«·ÿ·«»"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ÿ«„ ≈œ«—… „⁄·Ê„«  √⁄Ÿ«¡ ÂÌ∆… «· œ—Ì” Ê«·ÿ·«» ›Ì ﬁ”„ ⁄·Ê„ «·Õ«”»« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   6135
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "Main.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10905
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    Students.Show
End Sub

Private Sub Command2_Click()
    Me.Hide
    Lecturers.Show
End Sub

VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· «·œŒÊ·"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   " ”ÃÌ· «·œŒÊ·"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Œÿ« ›Ì „⁄·Ê„«   ”ÃÌ· «·œŒÊ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ·„… «·„—Ê—"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„” Œœ„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ”ÃÌ· «·œŒÊ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If (Text1 = "basim" And Text2 = "1234") Then
        Me.Hide
        Main.Show
    Else
        Label4.Visible = True
    End If
End Sub

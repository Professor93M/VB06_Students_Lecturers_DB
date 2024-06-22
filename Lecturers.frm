VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Lecturers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ãÚáæãÇÊ ÇáÇÓÇÊÐÉ"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÇáÕÝÍÉ ÇáÑÆíÓíÉ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9240
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Caption         =   "ÇáÈÍË"
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   12495
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Lecturers.frx":0000
         Left            =   4680
         List            =   "Lecturers.frx":001F
         TabIndex        =   35
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÈÍË"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         Height          =   405
         Left            =   7080
         TabIndex        =   22
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ãÚáæãÇÊ ÇáÇÓÊÇÐ"
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   12495
      Begin VB.Frame Frame4 
         Caption         =   "ÇáÊäÞá Èíä ÇáÈíÇäÇÊ"
         Height          =   975
         Left            =   360
         TabIndex        =   29
         Top             =   2520
         Width           =   3975
         Begin VB.CommandButton Command9 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3000
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2040
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command7 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1080
            TabIndex        =   31
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command6 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "ÇáÊÍßã ÈÇáÈíÇäÇÊ"
         Height          =   975
         Left            =   4440
         TabIndex        =   25
         Top             =   2520
         Width           =   4335
         Begin VB.CommandButton Command11 
            BackColor       =   &H00FFC0FF&
            Caption         =   "ÇáÊÞÑíÑ"
            Default         =   -1  'True
            Height          =   615
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00C0E0FF&
            Caption         =   "ÇäÔÇÁ"
            Height          =   615
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00C0C0FF&
            Caption         =   "ÍÐÝ"
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "ÊÍÏíË"
            Height          =   615
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ÇÖÇÝÉ"
            Height          =   615
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Gender"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Lecturers.frx":006E
         Left            =   9360
         List            =   "Lecturers.frx":0078
         TabIndex        =   19
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         DataField       =   "Salary"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         DataField       =   "Mobile"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         DataField       =   "Class"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6360
         TabIndex        =   13
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   9360
         TabIndex        =   11
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         DataField       =   "Job"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         DataField       =   "Age"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         DataField       =   "Department"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6360
         TabIndex        =   5
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "Fullname"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   9360
         TabIndex        =   3
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÌäÓ"
         Height          =   375
         Left            =   9360
         TabIndex        =   20
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÑÇÊÈ"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáåÇÊÝ"
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÏÑÌÉ ÇáæÙíÝíÉ"
         Height          =   375
         Left            =   6360
         TabIndex        =   14
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÓßä"
         Height          =   375
         Left            =   9360
         TabIndex        =   12
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáæÙíÝÉ"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÚãÑ"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÞÓã"
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÇÓã ÇáßÇãá"
         Height          =   375
         Left            =   9360
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Lecturers"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Lecturers.frx":0087
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "#"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Fullname"
         Caption         =   "ÇáÇÓã ÇáßÇãá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Department"
         Caption         =   "ÇáÞÓã"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Age"
         Caption         =   "ÇáÚãÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Job"
         Caption         =   "ÇáæÙíÝÉ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Address"
         Caption         =   "ÇáÓßä"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Class"
         Caption         =   "ÇáÏÑÌÉ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Mobile"
         Caption         =   "ÇáåÇÊÝ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Salary"
         Caption         =   "ÇáÑÇÊÈ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Gender"
         Caption         =   "ÇáÌäÓ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   900.284
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "äÙÇã ÅÏÇÑÉ ãÚáæãÇÊ ÃÚÙÇÁ åíÆÉ ÇáÊÏÑíÓ æÇáØáÇÈ Ýí ÞÓã Úáæã ÇáÍÇÓÈÇÊ"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Lecturers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim sql As String
    Adodc1.CommandType = adCmdText
    Field = Combo2.Text
    Select Case Field
        Case "ÇáÇÓã ÇáßÇãá"
            Field = "Fullname"
        Case "ÇáÞÓã"
            Field = "Department"
        Case "ÇáÚãÑ"
            Field = "Age"
        Case "ÇáæÙíÝÉ"
            Field = "Job"
        Case "ÇáÓßä"
            Field = "Address"
        Case "ÇáÏÑÌÉ"
            Field = "Class"
        Case "ÇáåÇÊÝ"
            Field = "Mobile"
        Case "ÇáÑÇÊÈ"
            Field = "Salary"
        Case Else
            Field = "Gender"
    End Select
    sql = "select * from lecturers where " & Field & " like '%" & Text9.Text & "%'"
    If (Text9 <> "") Then
        Adodc1.RecordSource = sql
        If Adodc1.Recordset.EOF Then
            MsgBox ("No Data")
        End If
    Else
        Adodc1.RecordSource = "select * from lecturers"
    End If
    Adodc1.Refresh
End Sub

Private Sub Command10_Click()
    On Error Resume Next
    Adodc1.Recordset.AddNew
End Sub

Private Sub Command11_Click()
    LECRep.Show
End Sub

Private Sub Command2_Click()
    Me.Hide
    Main.Show
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Adodc1.Recordset.Fields(1).Value = Text1.Text
    Adodc1.Recordset.Fields(2).Value = Text2.Text
    Adodc1.Recordset.Fields(3).Value = Text3.Text
    Adodc1.Recordset.Fields(4).Value = Text4.Text
    Adodc1.Recordset.Fields(5).Value = Text5.Text
    Adodc1.Recordset.Fields(6).Value = Text6.Text
    Adodc1.Recordset.Fields(7).Value = Text7.Text
    Adodc1.Recordset.Fields(8).Value = Text8.Text
    Adodc1.Recordset.Fields(9).Value = Combo1.Text
    Adodc1.Recordset.Update
    Adodc1.Refresh
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    Adodc1.Recordset.Fields(1).Value = Text1.Text
    Adodc1.Recordset.Fields(2).Value = Text2.Text
    Adodc1.Recordset.Fields(3).Value = Text3.Text
    Adodc1.Recordset.Fields(4).Value = Text4.Text
    Adodc1.Recordset.Fields(5).Value = Text5.Text
    Adodc1.Recordset.Fields(6).Value = Text6.Text
    Adodc1.Recordset.Fields(7).Value = Text7.Text
    Adodc1.Recordset.Fields(8).Value = Text8.Text
    Adodc1.Recordset.Fields(9).Value = Combo1.Text
    Adodc1.Recordset.Update
    Adodc1.Refresh
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    Adodc1.Recordset.Delete
End Sub

Private Sub Command6_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveLast
End Sub

Private Sub Command7_Click()
    On Error Resume Next
    Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command8_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveNext
End Sub

Private Sub Command9_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveFirst
End Sub

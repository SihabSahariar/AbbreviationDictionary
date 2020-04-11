VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abbreviation Dictionary By Sihab Sahariar"
   ClientHeight    =   3660
   ClientLeft      =   8595
   ClientTop       =   4320
   ClientWidth     =   8865
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   8865
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   3720
      Width           =   2535
   End
   Begin VB.ListBox List1 
      DataField       =   "ShortForm"
      DataSource      =   "Data1"
      Height          =   1230
      ItemData        =   "Form1.frx":6852
      Left            =   120
      List            =   "Form1.frx":759A
      TabIndex        =   7
      Top             =   2160
      Width           =   8655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GO"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      DataField       =   "FullForm"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      DataField       =   "ShortForm"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Dic"
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search  Abbreviation:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Short Form                           Full Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command3_Click()
 Dim content
    content = Trim(Text3.Text) & "*"
    content = "ShortForm like '" & content & "'"
    If Text3.Text <> "" Then
        Data1.Recordset.FindFirst content
    End If
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\Microsoft.net.dll"
Data1.RecordSource = "Table1"
End Sub


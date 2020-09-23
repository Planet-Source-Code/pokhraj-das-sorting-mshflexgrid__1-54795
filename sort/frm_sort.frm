VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "FlexGrid Sorting"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_desc 
      Caption         =   "Sort Desc"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmd_Asc 
      Caption         =   "Sort Asc"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmd_load 
      Caption         =   "Load File"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5530
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8415
   End
   Begin VB.Frame Frame2 
      Caption         =   " Select Option"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Sorting Demo By Pokhraj Das"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmd_Asc_Click()
Frame1.Caption = cmd_Asc.Caption
MSHFlexGrid1.Col = 1
For i = 1 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.RowSel = i
    MSHFlexGrid1.Sort = 3
Next
End Sub

Private Sub cmd_desc_Click()
Frame1.Caption = cmd_desc.Caption
MSHFlexGrid1.Col = 1
For i = 1 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.RowSel = i
    MSHFlexGrid1.Sort = 4
Next i
End Sub

Private Sub cmd_load_Click()
'for retrieving the records from database
rs.Open "select * from tab1", cn, adOpenDynamic, adLockBatchOptimistic
Set MSHFlexGrid1.DataSource = rs
Call flxresize
rs.Close
Set rs = Nothing
End Sub

Private Sub Form_Load()
'for connection using ado
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\P K TALAPATRA\Desktop\sort\dtn.mdb;Persist Security Info=False"
cn.Open
End Sub

Private Sub flxresize()
'adjusting grid
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 4500
MSHFlexGrid1.ColWidth(4) = 1290
MSHFlexGrid1.ColAlignment(4) = 3        'for align the data at center
End Sub

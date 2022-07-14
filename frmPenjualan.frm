VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPenjualan 
   Caption         =   "Penjualan"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   18090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Keluar"
      Height          =   435
      Left            =   16320
      TabIndex        =   15
      Top             =   720
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   17760
      Top             =   600
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   720
      Width           =   4935
   End
   Begin VB.TextBox TxtCabang 
      Height          =   405
      Left            =   14040
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox TxtKasir 
      Height          =   405
      Left            =   10920
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox TxtJam 
      Height          =   405
      Left            =   8280
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox TxtTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14040
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   7800
      Width           =   3735
   End
   Begin VB.TextBox TxtTgl 
      Height          =   405
      Left            =   5040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TxtFaktur 
      Height          =   405
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   17565
      _ExtentX        =   30983
      _ExtentY        =   11245
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Cari Obat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Cabang"
      Height          =   255
      Left            =   13080
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Kasir"
      Height          =   255
      Left            =   10320
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Jam"
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   6
      Top             =   7920
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "No. Faktur"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call bersih
End Sub

Private Sub Timer1_Timer()
    TxtJam.Text = Time
    TxtTgl.Text = Date
End Sub

Sub bersih()
    TxtFaktur.Enabled = False
    TxtTgl.Enabled = False
    TxtJam.Enabled = False
    TxtKasir.Enabled = False
    TxtCabang.Enabled = False
End Sub

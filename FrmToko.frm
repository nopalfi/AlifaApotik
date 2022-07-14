VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmToko 
   Caption         =   "Cabang"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Text            =   "Text5"
      Top             =   2400
      Width           =   4575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   390
      Left            =   240
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=alifa"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "alifa"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * From toko"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Keluar"
      Height          =   400
      Left            =   6000
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   4080
      TabIndex        =   11
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Rubah"
      Height          =   400
      Left            =   2280
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   400
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmToko.frx":0000
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "KodeToko"
         Caption         =   "KodeToko"
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
         DataField       =   "Nama_Toko"
         Caption         =   "Nama Toko"
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
         DataField       =   "Alamat"
         Caption         =   "Alamat"
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
         DataField       =   "HP"
         Caption         =   "HP"
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
         DataField       =   "Pengelola"
         Caption         =   "Pengelola"
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
         DataField       =   "KodeF"
         Caption         =   "Kode Faktur"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1920
      Width           =   5895
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   1440
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "FrmToko.frx":0015
      Top             =   720
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label6 
      Caption         =   "Kode Faktur"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Pimpinan "
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Handphone"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Alamat"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Toko"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmToko"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Or Text3.Text = "" Then
        MsgBox "Harap input data terlebih dahulu", vbOKOnly + vbInformation, "Informasi"
    Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!Nama_toko = Text1.Text
        Adodc1.Recordset!ALAMAT = Text2.Text
        Adodc1.Recordset!HP = Text3.Text
        Adodc1.Recordset!pengelola = Text4.Text
        Adodc1.Recordset!KodeF = Text5.Text
        Adodc1.Recordset.Update
        Call Bersih
    End If
End Sub

Private Sub Command2_Click()
    Adodc1.RecordSource = "Select * from toko where Kodetoko ='" & Label5.Caption & "'"
    Adodc1.Refresh
    If Not Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset!Nama_toko = Text1.Text
        Adodc1.Recordset!ALAMAT = Text2.Text
        Adodc1.Recordset!HP = Text3.Text
        Adodc1.Recordset!pengelola = Text4.Text
        Adodc1.Recordset!KodeF = Text5.Text
        Adodc1.Recordset.Update
        Call Bersih
    Else
        MsgBox "Data tidak ditemukan", vbOKOnly + vbInformation, "PERHATIAN"
    End If
End Sub

Private Sub Command3_Click()
    Adodc1.RecordSource = "Select * from toko where Kodetoko ='" & Label5.Caption & "'"
    Adodc1.Refresh
    If Not Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveFirst
        If MsgBox("Yakin akan menghapus TOKO " & _
          Text1.Text & "?", vbOKCancel) = vbOK Then
          Adodc1.Recordset.Delete
          Call Bersih
          MsgBox "Data berhasil di hapus", vbOKOnly + vbInformation, "PERHATIAN"
          DataGrid1.Refresh
        End If
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub DataGrid1_Click()
    If Not Adodc1.Recordset.EOF Then
        Label5.Caption = DataGrid1.Columns(0).Text
        Text1.Text = DataGrid1.Columns(1).Text
        Text2.Text = DataGrid1.Columns(2).Text
        Text3.Text = DataGrid1.Columns(3).Text
        Text4.Text = DataGrid1.Columns(4).Text
        Text5.Text = DataGrid1.Columns(5).Text
        Command2.Enabled = True
        Command3.Enabled = True
        Command1.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Call Bersih
End Sub

Sub Bersih()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Label5.Caption = ""
    Label5.Visible = False
    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = False
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub


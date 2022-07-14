VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmKategori 
   Caption         =   "Kategori"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rubah"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selesai / Keluar"
      Height          =   735
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DbGrid 
      Bindings        =   "FrmKategori.frx":0000
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "Kategori"
         Caption         =   "Kategori"
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
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   3495.118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   6120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "Select * From kategori"
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "FrmKategori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    frmMaster.Show (1)
    frmMaster.Adodc2.Refresh
    frmMaster.DataCombo3.Refresh
End Sub

Private Sub Command2_Click()
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields!kategori = Text1.Text
    Adodc1.Recordset.Update
    Call Bersih
End Sub

Private Sub Command3_Click()
    If Not Label1.Caption = "" Then
        Adodc1.RecordSource = "Select * from jenis where id ='" & Label1.Caption & "'"
        Adodc1.Refresh
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Fields!kategori = Text1.Text
        Adodc1.Recordset.Update
        Call Bersih
    End If
End Sub

Private Sub Command4_Click()
    If Not Label1.Caption = "" Then
        Adodc1.RecordSource = "Select * from jenis where id ='" & Label1.Caption & "'"
        Adodc1.Refresh
        If Not Adodc1.Recordset.EOF Then
            Adodc1.Recordset.MoveFirst
            If MsgBox("Yakin akan menghapus User = " & _
              Text1.Text & "?", vbOKCancel) = vbOK Then
              Adodc1.Recordset.Delete
              Call Bersih
              MsgBox "Data berhasil di hapus", vbOKOnly + vbInformation, "PERHATIAN"
            End If
        End If
    End If
End Sub

Private Sub DbGrid_Click()
    Label1.Caption = DbGrid.Columns(0)
    Text1.Text = DbGrid.Columns(1)
    Command3.Enabled = True
    Command4.Enabled = True
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Call Bersih
End Sub

Sub Bersih()
    Text1.Text = ""
    Label1.Caption = ""
    Label1.Visible = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Adodc1.Recordset.EOF Then
            Adodc1.RecordSource = "Select * from Kategori where kategori ='" & Text1.Text & "'"
            Adodc1.Refresh
            If Adodc1.Recordset.EOF Then
                If Label1.Caption = "" Then
                    Command2.Enabled = True
                    Command2.SetFocus
                End If
            Else
                Command3.Enabled = True
                Command4.Enabled = True
                Adodc1.Recordset.MoveFirst
                Label1.Caption = Adodc1.Recordset.Fields!id
            End If
        Else
            Command2.Enabled = True
            Command2.SetFocus
        End If
    End If
End Sub

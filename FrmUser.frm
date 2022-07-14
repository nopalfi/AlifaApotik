VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmUser 
   Caption         =   "Tambah User"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Rubah"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   3360
      Top             =   5880
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "DSN=alifaMariaDB"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "alifaMariaDB"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1800
      Top             =   5880
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
      Connect         =   "DSN=alifaMariaDB"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "alifaMariaDB"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select nama_toko from toko"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   5880
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
      RecordSource    =   "Select * From user"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmUser.frx":0000
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5106
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "USER"
         Caption         =   "USER"
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
         DataField       =   "PASSWORD"
         Caption         =   "PASSWORD"
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
         DataField       =   "TEMPAT"
         Caption         =   "TEMPAT"
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
         DataField       =   "HAK"
         Caption         =   "HAK"
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
            ColumnWidth     =   1739.906
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
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmUser.frx":0015
      Left            =   1200
      List            =   "FrmUser.frx":0022
      TabIndex        =   5
      Text            =   "Pilih Akses"
      Top             =   1560
      Width           =   2655
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FrmUser.frx":0043
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc2"
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nama_Toko"
      BoundColumn     =   "Nama_Toko"
      Text            =   "Pilih Cabang"
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Akses"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Cabang"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Not Text1.Text = "" Then
        Adodc1.RecordSource = "Select * from user where USER ='" & Text1.Text & "'"
        Adodc1.Refresh
        If Not Adodc1.Recordset.EOF Then
            MsgBox "User sudah ada harap gunakan nama user yang lain", vbInformation + vbOKOnly, "PERHATIAN"
        Else
            If Not Text2.Text = "" Then
                If Not DataCombo1.Text = "Pilih Cabang" Then
                    If Not Combo1.Text = "Pilih Akses" Then
                        Adodc1.Recordset.AddNew
                        Adodc1.Recordset!User = Text1.Text
                        Adodc1.Recordset!Password = Text2.Text
                        Adodc1.Recordset!Tempat = Label5.Caption
                        Adodc1.Recordset!HAK = Combo1.Text
                        Adodc1.Recordset.Update
                        Call Bersih
                    Else
                        MsgBox "Pilih Akses terlebih dahulu", vbOKOnly + vbInformation, "PERHATIAN"
                    End If
                Else
                    MsgBox "Pilih Cabang terlebih dahulu", vbOKOnly + vbInformation, "PERHATIAN"
                End If
            Else
                MsgBox "Isi Password terlebih dahulu", vbOKOnly + vbInformation, "PERHATIAN"
            End If
        End If
    Else
        MsgBox "Input User terlebih dahulu", vbOKOnly + vbInformation, "PERHATIAN"
    End If
End Sub


Private Sub Command2_Click()
    If Not Text1.Text = "" Then
        Adodc1.RecordSource = "Select * from user where USER ='" & Text1.Text & "'"
        Adodc1.Refresh
        If Adodc1.Recordset.RecordCount = 1 Then
            If Not Text2.Text = "" Then
                If Not DataCombo1.Text = "Pilih Cabang" Then
                    If Not Combo1.Text = "Pilih Akses" Then
                        Adodc1.Recordset.MoveFirst
                        Adodc1.Recordset!User = Text1.Text
                        Adodc1.Recordset!Password = Text2.Text
                        Adodc1.Recordset!Tempat = Label5.Caption
                        Adodc1.Recordset!HAK = Combo1.Text
                        Adodc1.Recordset.Update
                        Call Bersih
                    Else
                        MsgBox "Pilih Akses terlebih dahulu", vbOKOnly + vbInformation, "PERHATIAN"
                    End If
                Else
                    MsgBox "Pilih Cabang terlebih dahulu", vbOKOnly + vbInformation, "PERHATIAN"
                End If
            Else
                MsgBox "Isi Password terlebih dahulu", vbOKOnly + vbInformation, "PERHATIAN"
            End If
        
        Else
            MsgBox "User Tidak ditemukan", vbInformation + vbOKOnly, "Perhatian"
        End If
    Else
        MsgBox "Input User terlebih dahulu", vbOKOnly + vbInformation, "PERHATIAN"
    End If
End Sub

Private Sub Command3_Click()
    Adodc1.RecordSource = "Select * from user where USER ='" & Text1.Text & "'"
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
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    If Not DataCombo1.Text = "Pilih Cabang" Then
        Adodc3.RecordSource = "Select * From toko Where Nama_toko = '" & DataCombo1.Text & "'"
        Adodc3.Refresh
        Label5.Caption = Adodc3.Recordset!KodeToko
    End If
End Sub

Private Sub DataGrid1_DblClick()
    Text1.Text = DataGrid1.Columns(0)
    Text2.Text = DataGrid1.Columns(1)
    Adodc3.RecordSource = "Select * From toko Where KodeToko = '" & DataGrid1.Columns(2) & "'"
    Adodc3.Refresh
    If Not Adodc3.Recordset.BOF Then
        Adodc3.Recordset.MoveFirst
        DataCombo1.Text = Adodc3.Recordset!Nama_toko
        Label5.Caption = DataGrid1.Columns(2)
        Combo1.Text = DataGrid1.Columns(3)
        Text1.Enabled = False
        Command1.Enabled = False
        Command2.Enabled = True
        Command3.Enabled = True
    Else
        MsgBox "User tidak ditemukan", vbOKOnly + vbInformation, "PERHATIAN"
    End If
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Sub Bersih()
    Text1.Text = ""
    Text2.Text = ""
    DataCombo1.Text = "Pilih Cabang"
    Combo1.Text = "Pilih Akses"
    Command2.Enabled = False
    Command3.Enabled = False
    Label5.Visible = False
    Adodc1.RecordSource = "Select * from user"
    Adodc1.Refresh
End Sub


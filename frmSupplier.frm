VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   7200
      TabIndex        =   36
      Top             =   5040
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2640
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
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
      RecordSource    =   "countries"
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
   Begin MSDataListLib.DataCombo negaraCombo 
      Bindings        =   "frmSupplier.frx":0000
      DataSource      =   "Adodc2"
      Height          =   315
      Left            =   1680
      TabIndex        =   13
      Top             =   3840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "name"
      Text            =   "-- Negara --"
   End
   Begin VB.TextBox provinsiTxt 
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   3240
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   9960
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
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
      RecordSource    =   "select * from suplier"
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
      Bindings        =   "frmSupplier.frx":0015
      Height          =   3255
      Left            =   360
      TabIndex        =   34
      Top             =   6360
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5741
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "KODE"
         Caption         =   "KODE"
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
         DataField       =   "NAMA"
         Caption         =   "NAMA"
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
         DataField       =   "ALAMAT"
         Caption         =   "ALAMAT"
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
         DataField       =   "KOTA"
         Caption         =   "KOTA"
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
         DataField       =   "PROVINSI"
         Caption         =   "PROVINSI"
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
         DataField       =   "NEGARA"
         Caption         =   "NEGARA"
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
         DataField       =   "KODEPOS"
         Caption         =   "KODEPOS"
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
         DataField       =   "TELEPON"
         Caption         =   "TELEPON"
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
         DataField       =   "FAX"
         Caption         =   "FAX"
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
         DataField       =   "BANK"
         Caption         =   "BANK"
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
      BeginProperty Column10 
         DataField       =   "ACC"
         Caption         =   "ACC"
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
      BeginProperty Column11 
         DataField       =   "ATASNAMA"
         Caption         =   "ATASNAMA"
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
      BeginProperty Column12 
         DataField       =   "KONTAK"
         Caption         =   "KONTAK"
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
      BeginProperty Column13 
         DataField       =   "EMAIL"
         Caption         =   "EMAIL"
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
      BeginProperty Column14 
         DataField       =   "KETERANGAN"
         Caption         =   "KETERANGAN"
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
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton batalCmd 
      Caption         =   "Batal"
      Height          =   495
      Left            =   8640
      TabIndex        =   25
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton ubahCmd 
      Caption         =   "Ubah"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   24
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton simpanCmd 
      Caption         =   "Simpan"
      Default         =   -1  'True
      Height          =   495
      Left            =   4320
      TabIndex        =   23
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox keteranganTxt 
      Height          =   405
      Left            =   6480
      TabIndex        =   22
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox emailTxt 
      Height          =   405
      Left            =   6480
      TabIndex        =   21
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox kontakTxt 
      Height          =   405
      Left            =   6480
      TabIndex        =   20
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox atasNamaTxt 
      Height          =   405
      Left            =   6480
      TabIndex        =   19
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox accTxt 
      Height          =   405
      Left            =   6480
      TabIndex        =   18
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox bankTxt 
      Height          =   405
      Left            =   6480
      TabIndex        =   17
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox faxTxt 
      Height          =   405
      Left            =   6480
      TabIndex        =   16
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox teleponTxt 
      Height          =   405
      Left            =   1680
      TabIndex        =   15
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox kodePostTxt 
      Height          =   405
      Left            =   1680
      TabIndex        =   14
      Top             =   4440
      Width           =   3135
   End
   Begin VB.TextBox kotaTxt 
      Height          =   405
      Left            =   1680
      TabIndex        =   11
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox alamatTxt 
      Height          =   645
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox namaTxt 
      Height          =   405
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox kodeTxt 
      Height          =   405
      Left            =   1680
      TabIndex        =   8
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "(*) Opsional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   35
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Daftar Supplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   33
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label15 
      Caption         =   "Keterangan (*)"
      Height          =   375
      Left            =   5160
      TabIndex        =   32
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "Email"
      Height          =   375
      Left            =   5160
      TabIndex        =   31
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Kontak"
      Height          =   375
      Left            =   5160
      TabIndex        =   30
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Atas Nama"
      Height          =   375
      Left            =   5160
      TabIndex        =   29
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "ACC (*)"
      Height          =   375
      Left            =   5160
      TabIndex        =   28
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Bank (*)"
      Height          =   375
      Left            =   5160
      TabIndex        =   27
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Fax (*)"
      Height          =   255
      Left            =   5160
      TabIndex        =   26
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Telepon"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Kode Pos"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Negara"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Provinsi"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Kota"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Kode"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub batalCmd_Click()
    Unload Me
    frmMaster.Show (1)
    frmMaster.Adodc1.Refresh
    frmMaster.DataCombo1.Refresh
End Sub

Private Sub Command1_Click()
Adodc1.RecordSource = "Select * from suplier where KODE ='" & kodeTxt.Text & "'"
    Adodc1.Refresh
    If Not Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveFirst
        If MsgBox("Yakin akan menghapus Supplier " & kodeTxt.Text & "?", vbOKCancel) = vbOK Then
          Adodc1.Recordset.Delete
          Call clear
          MsgBox "Data berhasil di hapus", vbOKOnly + vbInformation, "PERHATIAN"
          DataGrid1.Refresh
        End If
    End If
End Sub

Private Sub DataGrid1_Click()
    If Not Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveFirst
        kodeTxt.Text = DataGrid1.Columns(0).Text
        namaTxt.Text = DataGrid1.Columns(1).Text
        alamatTxt.Text = DataGrid1.Columns(2).Text
        kotaTxt.Text = DataGrid1.Columns(3).Text
        provinsiTxt.Text = DataGrid1.Columns(4).Text
        negaraCombo.Text = DataGrid1.Columns(5).Text
        kodePostTxt.Text = DataGrid1.Columns(6).Text
        teleponTxt.Text = DataGrid1.Columns(7).Text
        faxTxt.Text = DataGrid1.Columns(8).Text
        bankTxt.Text = DataGrid1.Columns(9).Text
        accTxt.Text = DataGrid1.Columns(10).Text
        atasNamaTxt.Text = DataGrid1.Columns(11).Text
        kontakTxt.Text = DataGrid1.Columns(12).Text
        emailTxt.Text = DataGrid1.Columns(13).Text
        keteranganTxt.Text = DataGrid1.Columns(14).Text
        ubahCmd.Enabled = True
    End If
End Sub

Private Sub simpanCmd_Click()
If Not kodeTxt.Text = "" Then
    If Not namaTxt.Text = "" Then
        If Not alamatTxt.Text = "" Then
            If Not kotaTxt.Text = "" Then
                If Not provinsiTxt.Text = "" Then
                    If Not negaraCombo.Text = "" Then
                        If Not kodePostTxt.Text = "" Then
                            If Not teleponTxt.Text = "" Then
                                If Not atasNamaTxt.Text = "" Then
                                    If Not kontakTxt.Text = "" Then
                                        If Not emailTxt.Text = "" Then
                                            simpanCmd.Enabled = True
                                            Adodc1.Recordset.AddNew
                                            Adodc1.Recordset!KODE = kodeTxt.Text
                                            Adodc1.Recordset!NAMA = namaTxt.Text
                                            Adodc1.Recordset!ALAMAT = alamatTxt.Text
                                            Adodc1.Recordset!KOTA = kotaTxt.Text
                                            Adodc1.Recordset!PROVINSI = provinsiTxt.Text
                                            Adodc1.Recordset!NEGARA = negaraCombo.Text
                                            Adodc1.Recordset!KODEPOS = kodePostTxt.Text
                                            Adodc1.Recordset!TELEPON = teleponTxt.Text
                                            Adodc1.Recordset!FAX = faxTxt.Text
                                            Adodc1.Recordset!BANK = bankTxt.Text
                                            Adodc1.Recordset!ACC = accTxt.Text
                                            Adodc1.Recordset!ATASNAMA = atasNamaTxt.Text
                                            Adodc1.Recordset!KONTAK = kontakTxt.Text
                                            Adodc1.Recordset!EMAIL = emailTxt.Text
                                            Adodc1.Recordset!KETERANGAN = keteranganTxt.Text
                                            Adodc1.Recordset.Update
                                            MsgBox "Cabang baru berhasil ditambahkan!", vbOKOnly + vbExclamation, "SUKSES"
                                            Call clear
                                        Else
                                            MsgBox "Masukkan Email Terlebih dahulu!", vbOKOnly + vbInformation, "PERHATIAN"
                                        End If
                                    Else
                                        MsgBox "Masukkan nomor kontak terlebih dahulu!", vbOKOnly + vbInformation, "PERHATIAN"
                                    End If
                                Else
                                    MsgBox "Masukkan Info Atas Nama!", vbOKOnly + vbInformation, "PERHATIAN"
                                End If
                            Else
                                MsgBox "Masukkan Nomor Telepon Terlebih Dahulu!", vbOKOnly + vbInformation, "PERHATIAN"
                            End If
                        Else
                            MsgBox "Masukkan Kode Pos Terlebih daahulu!", vbOKOnly + vbInformation, "PERHATIAN"
                        End If
                    Else
                        MsgBox "Pilih salah satu negara", vbOKOnly + vbInformation, "PERHATIAN"
                    End If
                Else
                    MsgBox "Masukkan Provinsi terlebih dahulu!", vbOKOnly + vbInformation, "PERHATIAN"
                End If
            Else
                MsgBox "Masukkan Kota Terlebih Dahulu!", vbOKOnly + vbInformation, "PERHATIAN"
            End If
        Else
            MsgBox "Masukkan Alamat Terlebih dahulu!", vbOKOnly + vbInformation, "PERHATIAN"
        End If
    Else
        MsgBox "Masukkan Nama Toko terlebih dahulu!", vbOKOnly + vbInformation, "PERHATIAN"
    End If
Else
    MsgBox "Masukkan Kode Toko Terlebih dahulu!", vbOKOnly + vbInformation, "PERHATIAN"
End If
End Sub

Private Sub clear()
    kodeTxt.Text = ""
    namaTxt.Text = ""
    alamatTxt.Text = ""
    kotaTxt.Text = ""
    provinsiTxt.Text = ""
    negaraCombo.Text = ""
    kodePostTxt.Text = ""
    teleponTxt.Text = ""
    faxTxt.Text = ""
    bankTxt.Text = ""
    accTxt.Text = ""
    atasNamaTxt.Text = ""
    kontakTxt.Text = ""
    emailTxt.Text = ""
    keteranganTxt.Text = ""
    kodeTxt.SetFocus
End Sub

Private Sub ubahCmd_Click()
    Adodc1.RecordSource = "select * from suplier where kode ='" & kodeTxt.Text & "'"
    Adodc1.Refresh
    If Not Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset!KODE = kodeTxt.Text
        Adodc1.Recordset!NAMA = namaTxt.Text
        Adodc1.Recordset!ALAMAT = alamatTxt.Text
        Adodc1.Recordset!KOTA = kotaTxt.Text
        Adodc1.Recordset!PROVINSI = provinsiTxt.Text
        Adodc1.Recordset!NEGARA = negaraCombo.Text
        Adodc1.Recordset!KODEPOS = kodePostTxt.Text
        Adodc1.Recordset!TELEPON = teleponTxt.Text
        Adodc1.Recordset!FAX = faxTxt.Text
        Adodc1.Recordset!BANK = bankTxt.Text
        Adodc1.Recordset!ACC = accTxt.Text
        Adodc1.Recordset!ATASNAMA = atasNamaTxt.Text
        Adodc1.Recordset!KONTAK = kontakTxt.Text
        Adodc1.Recordset!EMAIL = emailTxt.Text
        Adodc1.Recordset!KETERANGAN = keteranganTxt.Text
        Adodc1.Recordset.Update
        MsgBox "Data berhasil dirubah", vbOKOnly + vbInformation, "SUKSES"
        Call clear
    Else
        MsgBox "Data tidak ditemukan", vbOKOnly + vbInformation, "PERHATIAN"
    End If
End Sub

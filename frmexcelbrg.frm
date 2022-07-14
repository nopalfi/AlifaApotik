VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmexcelbrg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excel Barang"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   17640
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   9000
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
      RecordSource    =   "masterbarang"
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
      Caption         =   "Keluar"
      Height          =   615
      Left            =   15600
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Simpan"
      Height          =   615
      Left            =   13800
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox lstTable 
      Height          =   7275
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtfile 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tampilkan Ke Tabel"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cari File Excel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   9360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid grdExcel 
      Height          =   7275
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   12832
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
End
Attribute VB_Name = "frmexcelbrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xConn As ADODB.Connection
Dim rsTable As ADODB.Recordset
Dim rsExcel As ADODB.Recordset
Dim strExcel As String
Dim x As Long

Private Sub Open_Excel(FilePath As String)
On Error GoTo err_Handler

    Set xConn = New ADODB.Connection
    With xConn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & FilePath & _
                            ";Extended Properties=Excel 8.0;"
        .Open
    End With

Exit Sub
err_Handler:
    MsgBox "Error on Open_Excel :" & Err.Number & " " & Err.Description
End Sub

Private Sub List_Table()
On Error GoTo err_Handler

    lstTable.clear
    Open_Excel Me.txtfile.Text
    Set rsTable = xConn.OpenSchema(adSchemaTables)
    Do While Not rsTable.EOF
    lstTable.AddItem rsTable.Fields("TABLE_NAME").Value
    rsTable.MoveNext
    Loop
    Set rsTable = Nothing
    xConn.Close
    If Not Me.lstTable.ListCount = 0 Then
       Me.lstTable.ListIndex = 0
       Data_Excel Me.lstTable.Text
    End If

Exit Sub
err_Handler:
    MsgBox "Error on Open_Excel :" & Err.Number & " " & Err.Description
End Sub

Private Sub Data_Excel(Sheet As String)
Open_Excel Me.txtfile.Text

Set rsExcel = New ADODB.Recordset
strExcel = "SELECT * FROM [" & Sheet & "]"

With rsExcel
    .CursorLocation = adUseClient
    .Open strExcel, xConn, adOpenKeyset, adLockReadOnly
    .ActiveConnection = Nothing
End With
Set Me.grdExcel.DataSource = rsExcel
xConn.Close
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Command1_Click()
    dlgOpen.Filter = "Excel Files (*.xls)|*.xls"
    dlgOpen.ShowOpen
    Me.txtfile.Text = dlgOpen.FileName
    Command2.Enabled = True
End Sub

'Untuk load data excel, me list table/sheet, dan menampilkan di datagrid.

Private Sub Command2_Click()
    If txtfile.Text = "" Then _
       MsgBox "Pilih file excel untuk proses import": Exit Sub
    If LCase(Right$(txtfile.Text, 4)) <> ".xls" Then _
       MsgBox "File harus dalam format Excel(.xls)": Exit Sub
    List_Table
    Command2.Enabled = False
End Sub

Private Sub Command3_Click()
    If Not rsExcel.EOF Then
        rsExcel.MoveFirst
        x = 0
        Do Until rsExcel.EOF
          Adodc1.Recordset.AddNew
          Adodc1.Recordset!KODEITEM = rsExcel!KODEITEM
          Adodc1.Recordset!BARCODE = rsExcel!BARCODE
          Adodc1.Recordset!NAMAITEM = rsExcel!NAMAITEM
          Adodc1.Recordset!jenis = rsExcel!jenis
          Adodc1.Recordset!satuan = rsExcel!satuan
          Adodc1.Recordset!kategori = rsExcel!kategori
          Adodc1.Recordset!merek = rsExcel!merek
          Adodc1.Recordset!SATUANDASAR = rsExcel!SATUANDASAR
          Adodc1.Recordset!STOKMINIMUM = rsExcel!STOKMINIMUM
          Adodc1.Recordset!BERATSATUAN = rsExcel!BERATSATUANTERKECIL
          Adodc1.Recordset!HARGAPOKOK = rsExcel!HARGAPOKOK
          Adodc1.Recordset!HARGAJUAL = rsExcel!HARGAJUAL
          Adodc1.Recordset!KODESUPLIER = rsExcel!KODESUPPLIER
          Adodc1.Recordset.Update
          x = x + 1
          rsExcel.MoveNext
        Loop
     End If
     MsgBox "Data berhasil di upload = " & x & " dari " & rsExcel.RecordCount, vbOKOnly, "INFORMASI"
End Sub

Private Sub Command4_Click()
    Unload Me
    frmMaster.Show (1)
End Sub

Private Sub Form_Load()
    Command2.Enabled = False
    txtfile.Text = ""
    Adodc1.Refresh
End Sub

VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMaster 
   Caption         =   "Master Barang"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   375
      Left            =   8880
      Top             =   7560
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
      RecordSource    =   "select * from masterbarang"
      Caption         =   "Adodc7"
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
      Bindings        =   "frmMaster.frx":0000
      Height          =   3255
      Left            =   120
      TabIndex        =   36
      Top             =   4200
      Width           =   10095
      _ExtentX        =   17806
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "KODEITEM"
         Caption         =   "KODEITEM"
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
         DataField       =   "BARCODE"
         Caption         =   "BARCODE"
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
         DataField       =   "NAMAITEM"
         Caption         =   "NAMAITEM"
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
         DataField       =   "JENIS"
         Caption         =   "JENIS"
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
         DataField       =   "SATUAN"
         Caption         =   "SATUAN"
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
         DataField       =   "KATEGORI"
         Caption         =   "KATEGORI"
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
         DataField       =   "MEREK"
         Caption         =   "MEREK"
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
         DataField       =   "SATUANDASAR"
         Caption         =   "SATUANDASAR"
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
         DataField       =   "STOKMINIMUM"
         Caption         =   "STOKMINIMUM"
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
         DataField       =   "BERATSATUAN"
         Caption         =   "BERATSATUAN"
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
         DataField       =   "HARGAPOKOK"
         Caption         =   "HARGAPOKOK"
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
         DataField       =   "HARGAJUAL"
         Caption         =   "HARGAJUAL"
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
         DataField       =   "KODESUPLIER"
         Caption         =   "KODESUPLIER"
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
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   375
      Left            =   7440
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      RecordSource    =   "suplier"
      Caption         =   "Adodc6"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   6120
      Top             =   7560
      Visible         =   0   'False
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
      RecordSource    =   "satuan_dasar"
      Caption         =   "Adodc5"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4680
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      RecordSource    =   "merek"
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   3240
      Top             =   7560
      Visible         =   0   'False
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
      RecordSource    =   "satuan"
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
      Top             =   7560
      Visible         =   0   'False
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
      RecordSource    =   "kategori"
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
      Left            =   360
      Top             =   7560
      Visible         =   0   'False
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
      RecordSource    =   "jenis"
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
   Begin VB.CommandButton Command10 
      Caption         =   "Tambah Barang Dari Excel"
      Height          =   495
      Left            =   5520
      TabIndex        =   35
      Top             =   3600
      Width           =   4095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   8400
      TabIndex        =   34
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Rubah"
      Height          =   375
      Left            =   6960
      TabIndex        =   33
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   5520
      TabIndex        =   32
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Height          =   495
      Left            =   9480
      Picture         =   "frmMaster.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2520
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Bindings        =   "frmMaster.frx":05C1
      Height          =   315
      Left            =   7080
      TabIndex        =   30
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "NAMA"
      Text            =   "DataCombo6"
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   7080
      TabIndex        =   29
      Text            =   "Text7"
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   7080
      TabIndex        =   28
      Text            =   "Text6"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   7080
      TabIndex        =   27
      Text            =   "Text5"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7080
      TabIndex        =   26
      Text            =   "Text4"
      Top             =   600
      Width           =   2895
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "frmMaster.frx":05D6
      Height          =   315
      Left            =   7080
      TabIndex        =   25
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Satuan"
      Text            =   "DataCombo5"
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   9360
      Picture         =   "frmMaster.frx":05EB
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   0
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "frmMaster.frx":0B97
      Height          =   315
      Left            =   1440
      TabIndex        =   23
      Top             =   3720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Merek"
      Text            =   "DataCombo4"
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   4320
      Picture         =   "frmMaster.frx":0BAC
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   4320
      Picture         =   "frmMaster.frx":1158
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3000
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frmMaster.frx":1704
      Height          =   315
      Left            =   1440
      TabIndex        =   20
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Kategori"
      Text            =   "DataCombo3"
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   4320
      Picture         =   "frmMaster.frx":1719
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2400
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frmMaster.frx":1CC5
      Height          =   315
      Left            =   1440
      TabIndex        =   18
      Top             =   2520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Satuan"
      Text            =   "DataCombo2"
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   4320
      Picture         =   "frmMaster.frx":1CDA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmMaster.frx":2286
      Height          =   315
      Left            =   1440
      TabIndex        =   16
      Top             =   1920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Jenis"
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "frmMaster.frx":229B
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1440
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1440
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label13 
      Caption         =   "Suplier"
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Harga Jual"
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Harga Pokok"
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Berat Satuan"
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Stok Minimum"
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Satuan Dasar"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Merek"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Kategori"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Satuan"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Jenis"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Barang"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Barcode"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Item"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    FrmJenis.Show (1)
End Sub

Private Sub Command10_Click()
    Unload Me
    frmexcelbrg.Show
End Sub

Private Sub Command2_Click()
    Unload Me
    FrmSatuan.Show (1)
End Sub

Private Sub Command3_Click()
    Unload Me
    FrmKategori.Show (1)
End Sub

Private Sub Command4_Click()
    Unload Me
    Frmmerek.Show (1)
End Sub

Private Sub Command5_Click()
Unload Me
FrmSatuanD.Show (1)
End Sub

Private Sub Command6_Click()
Unload Me
frmSupplier.Show (1)
End Sub

Private Sub Command9_Click()
    Unload Me
End Sub

Private Sub DataList1_Click()

End Sub

Private Sub DataGrid1_Click()
    If Not Adodc7.Recordset.EOF Then
        Text1.Text = DataGrid1.Columns(0).Text
        Text2.Text = DataGrid1.Columns(1).Text
        Text3.Text = DataGrid1.Columns(2).Text
        DataCombo1.Text = DataGrid1.Columns(3).Text
        DataCombo2.Text = DataGrid1.Columns(4).Text
        DataCombo3.Text = DataGrid1.Columns(5).Text
        DataCombo4.Text = DataGrid1.Columns(6).Text
        DataCombo5.Text = DataGrid1.Columns(7).Text
        Text4.Text = DataGrid1.Columns(8).Text
        Text5.Text = DataGrid1.Columns(9).Text
        Text6.Text = DataGrid1.Columns(10).Text
        Text7.Text = DataGrid1.Columns(11).Text
        DataCombo6.Text = DataGrid1.Columns(12).Text
    End If
End Sub

Private Sub Form_Load()
    Call Bersih
End Sub

Sub Bersih()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = "0"
    Text5.Text = "0"
    Text6.Text = "0"
    Text7.Text = "0"
    DataCombo1.Text = "Pilih Jenis"
    DataCombo2.Text = "Pilih Satuan"
    DataCombo3.Text = "Kategori Obat"
    DataCombo4.Text = "Merek Obat"
    DataCombo5.Text = "Satuan Dasar"
    DataCombo6.Text = "Suplier Obat"
End Sub

VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmBeg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alifa Farma "
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   12585
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3675
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   10890
      Left            =   0
      Picture         =   "FrmBeg.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17490
   End
   Begin VB.Menu mnPenjualan 
      Caption         =   "&Penjualan"
   End
   Begin VB.Menu mnPembelian 
      Caption         =   "Pe&mbelian"
   End
   Begin VB.Menu mnSetting 
      Caption         =   "&Setting"
      Begin VB.Menu mnMaster 
         Caption         =   "&Master Barang"
      End
      Begin VB.Menu mnToko 
         Caption         =   "&Input Cabang"
      End
      Begin VB.Menu mnUser 
         Caption         =   "Input &User"
      End
   End
   Begin VB.Menu mnKeluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "FrmBeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Image1.Width = FrmBeg.Width
    Image1.Height = StatusBar1.Top
    If StatusBar1.Panels(2) = "Kasir" Then
        mnPenjualan.Visible = True
        mnPembelian.Visible = False
        mnSetting.Visible = False
    ElseIf StatusBar1.Panels(2) = "Administrator" Then
        mnPenjualan.Visible = False
        mnPembelian.Visible = True
        mnSetting.Visible = True
    ElseIf StatusBar1.Panels(2) = "Admin" Then
        mnPenjualan.Visible = True
        mnPembelian.Visible = True
        mnSetting.Visible = False
    End If
End Sub

Private Sub mnKeluar_Click()
    End
End Sub

Private Sub mnMaster_Click()
    frmMaster.Show (1)
End Sub

Private Sub mnToko_Click()
    FrmToko.Show (1)
End Sub

Private Sub mnUser_Click()
    FrmUser.Show (1)
End Sub

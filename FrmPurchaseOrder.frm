VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPurchaseOrder 
   BackColor       =   &H0000FF00&
   Caption         =   "Purchase Order"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   15465
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnHitung 
      Caption         =   "Hitung"
      Height          =   495
      Left            =   11400
      TabIndex        =   34
      Top             =   960
      Width           =   1500
   End
   Begin VB.TextBox txtNoPo 
      Height          =   495
      Left            =   1800
      TabIndex        =   33
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Frame FrameHitung 
      BackColor       =   &H0000FF00&
      Height          =   2415
      Left            =   7680
      TabIndex        =   10
      Top             =   1800
      Width           =   7455
      Begin VB.TextBox txtPpnInvoice 
         Height          =   495
         Left            =   1680
         TabIndex        =   16
         Top             =   960
         Width           =   2000
      End
      Begin VB.TextBox txtTotalInvoice 
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   1680
         Width           =   2000
      End
      Begin VB.TextBox txtInvoice 
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   2000
      End
      Begin VB.TextBox txtRetensi 
         Height          =   495
         Left            =   5280
         TabIndex        =   13
         Top             =   240
         Width           =   2000
      End
      Begin VB.TextBox txtTotalRetensi 
         Height          =   495
         Left            =   5280
         TabIndex        =   12
         Top             =   1680
         Width           =   2000
      End
      Begin VB.TextBox txtPpnRetensi 
         Height          =   495
         Left            =   5280
         TabIndex        =   11
         Top             =   960
         Width           =   2000
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Total Invoice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         Caption         =   "PPN Invoice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000FF00&
         Caption         =   "Invoice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "Retensi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H0000FF00&
         Caption         =   "PPN Retensi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H0000FF00&
         Caption         =   "Total Retensi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   17
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.TextBox txtHargaPo 
      Height          =   495
      Left            =   9240
      TabIndex        =   9
      Top             =   960
      Width           =   2000
   End
   Begin VB.ComboBox cmbJenisInvoice 
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   3480
      Width           =   2445
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton btnKeluar 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   1500
   End
   Begin VB.TextBox txtNamaProject 
      Height          =   975
      Left            =   1800
      TabIndex        =   3
      Top             =   1800
      Width           =   2600
   End
   Begin VB.TextBox txtNo 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbPerusahaan 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   3960
      Width           =   1245
   End
   Begin VB.TextBox txtDeskripsi 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2880
      Width           =   5715
   End
   Begin MSAdodcLib.Adodc adoPo 
      Height          =   495
      Left            =   9240
      Top             =   6360
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
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
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Worka\Aplikasi Master\db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Worka\Aplikasi Master\db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from purchaseOrder"
      Caption         =   "Data Purchase Order"
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
   Begin MSComCtl2.DTPicker dtpTanggalPo 
      Height          =   375
      Left            =   9240
      TabIndex        =   23
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   99549185
      CurrentDate     =   41512
   End
   Begin MSDataGridLib.DataGrid dtgPurchaseOrder 
      Bindings        =   "FrmPurchaseOrder.frx":0000
      Height          =   3135
      Left            =   360
      TabIndex        =   24
      Top             =   5400
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   5530
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "no"
         Caption         =   "no"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nomorPo"
         Caption         =   "nomorPo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "namaProject"
         Caption         =   "namaProject"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "deskripsi"
         Caption         =   "deskripsi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "tanggalPo"
         Caption         =   "tanggalPo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "jenisInvoice"
         Caption         =   "jenisInvoice"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "tanggalInvoice"
         Caption         =   "tanggalInvoice"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "hargaPo"
         Caption         =   "hargaPo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "invoice"
         Caption         =   "invoice"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "ppnInvoice"
         Caption         =   "ppnInvoice"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "totalInvoice"
         Caption         =   "totalInvoice"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "invoiceRetensi"
         Caption         =   "invoiceRetensi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "retensi"
         Caption         =   "retensi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "ppnRetensi"
         Caption         =   "ppnRetensi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "totalRetensi"
         Caption         =   "totalRetensi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "perusahaan"
         Caption         =   "perusahaan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   10920
      Picture         =   "FrmPurchaseOrder.frx":0014
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   3540
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Harga PO"
      Height          =   495
      Left            =   7800
      TabIndex        =   32
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FF00&
      Caption         =   "No P.O"
      Height          =   495
      Left            =   360
      TabIndex        =   31
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "Deskripsi"
      Height          =   495
      Left            =   360
      TabIndex        =   30
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "Tanggal PO"
      Height          =   495
      Left            =   7800
      TabIndex        =   29
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Caption         =   "Jenis Invoice"
      Height          =   495
      Left            =   360
      TabIndex        =   28
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Nama Project"
      Height          =   495
      Left            =   360
      TabIndex        =   27
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000FF00&
      Caption         =   "No"
      Height          =   495
      Left            =   360
      TabIndex        =   26
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000FF00&
      Caption         =   "Perusahaan"
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
    For Each x In Me
        If TypeOf x Is TextBox Then
            x.Text = ""
        End If
    Next
End Sub

Sub aktifTxt()
    For Each x In Me
        If TypeOf x Is TextBox Then
            x.Enabled = True
        End If
    Next
    cmbJenisInvoice.Enabled = True
    cmbPerusahaan.Enabled = True
End Sub

Sub nonAktifTxt()
    For Each x In Me
        If TypeOf x Is TextBox Then
            x.Enabled = False
        End If
    Next
    cmbJenisInvoice.Enabled = False
    cmbPerusahaan.Enabled = False
End Sub

Sub aktifBtn()
    For Each x In Me
        If TypeOf x Is CommandButton Then
            x.Enabled = True
        End If
    Next
End Sub

Sub nonAktifBtn()
    For Each x In Me
        If TypeOf x Is CommandButton Then
            x.Enabled = False
        End If
    Next
    btnTambah.Enabled = True
    btnKeluar.Enabled = True
End Sub

Sub isiJenisInvoice()
    cmbJenisInvoice.Text = "-Pilih-"
    cmbJenisInvoice.AddItem "Pre Survey"
    cmbJenisInvoice.AddItem "Installation"
    cmbJenisInvoice.AddItem "NMS"
    cmbJenisInvoice.AddItem "Validation"
    cmbJenisInvoice.AddItem "Dismantle"
    cmbJenisInvoice.AddItem "Turn Key"
    cmbJenisInvoice.AddItem "Cut Over & Dismantle"
    cmbPerusahaan.Text = "-Pilih-"
    cmbPerusahaan.AddItem "DMM"
    cmbPerusahaan.AddItem "DTI"
End Sub

Sub simpan()
    On Error Resume Next
    With adoPo.Recordset
        .AddNew
        .Fields(0) = txtNo.Text
        .Fields(1) = txtNoPo.Text
        .Fields(2) = txtNamaProject.Text
        .Fields(3) = txtDeskripsi.Text
        .Fields(4) = dtpTanggalPo.Value
        .Fields(5) = cmbJenisInvoice.Text
        .Fields(6) = ""
        .Fields(7) = txtHargaPo.Text
        .Fields(8) = txtInvoice.Text
        .Fields(9) = txtPpnInvoice.Text
        .Fields(10) = txtTotalInvoice.Text
        .Fields(11) = ""
        .Fields(12) = txtRetensi.Text
        .Fields(13) = txtPpnRetensi.Text
        .Fields(14) = txtTotalRetensi
        .Fields(15) = cmbPerusahaan.Text
        .Update
    End With
    dtgPurchaseOrder.Refresh
    On Error GoTo 0
End Sub

Sub hapus()
    On Error Resume Next
    If adoPo.Recordset.RecordCount = 0 Then
        MsgBox "Data sudah tidak ada", , "Hapus"
    Else
        If MsgBox("Ingin hapus data ini?", vbCritical + vbYesNo, "Peringatan") = vbYes Then
            adoPo.Recordset.Delete
        Else
            Exit Sub
        End If
    End If
    On Error GoTo 0
End Sub

Sub noAutomatis()
    Dim nomor As String
    adoPo.Recordset.Sort = "no"
    adoPo.RecordSource = "select * from purchaseOrder"
    If adoPo.Recordset.RecordCount = 0 Then
        nomor = 1
    Else
        adoPo.Recordset.MoveLast
        nomor = adoPo.Recordset.Fields(0) + 1
    End If
    txtNo.Text = nomor
End Sub

Private Sub btnHapus_Click()
    hapus
    noAutomatis
End Sub

Private Sub btnHitung_Click()
    txtInvoice.Text = Val(txtHargaPo.Text) * 90 / 100
    txtPpnInvoice.Text = Val(txtInvoice.Text) * 10 / 100
    txtTotalInvoice.Text = Val(txtInvoice.Text) + Val(txtPpnInvoice.Text)
    txtRetensi.Text = Val(txtHargaPo.Text) * 10 / 100
    txtPpnRetensi.Text = Val(txtRetensi.Text) * 10 / 100
    txtTotalRetensi.Text = Val(txtRetensi.Text) + Val(txtPpnRetensi.Text)
End Sub

Private Sub btnKeluar_Click()
    If MsgBox("Anda ingin keluar?", vbQuestion + vbYesNo, "Keluar") = vbYes Then
        End
    End If
End Sub

Private Sub btnSimpan_Click()
    simpan
    noAutomatis
End Sub

Private Sub btnTambah_Click()
    bersih
    noAutomatis
    aktifBtn
    aktifTxt
    txtHargaPo.Text = 0
    txtNo.SetFocus
End Sub

Private Sub cmbJenisInvoice_Click()
    If cmbJenisInvoice.Text = "Installation" Or cmbJenisInvoice.Text = "Turn Key" Then
        FrameHitung.Visible = True
        btnHitung.Visible = True
    Else
        FrameHitung.Visible = False
        btnHitung.Visible = False
    End If
End Sub

Private Sub Form_Activate()
    bersih
    nonAktifBtn
    nonAktifTxt
    FrameHitung.Visible = False
    btnHitung.Visible = False
    isiJenisInvoice
End Sub

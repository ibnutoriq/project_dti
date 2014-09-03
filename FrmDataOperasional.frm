VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmDataOperasional 
   BackColor       =   &H0000FF00&
   Caption         =   "Data Operasional"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   13380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnUbah 
      Caption         =   "Ubah"
      Height          =   495
      Left            =   4200
      TabIndex        =   23
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton btnHitungSaldo 
      Caption         =   "Tambahkan"
      Height          =   495
      Left            =   11400
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton btnHitungBk 
      Caption         =   "Hitung"
      Height          =   495
      Left            =   11400
      TabIndex        =   22
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   21
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtJumlahSaldo 
      Height          =   495
      Left            =   7800
      TabIndex        =   6
      Top             =   240
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Height          =   855
      Left            =   6000
      TabIndex        =   18
      Top             =   1680
      Width           =   5175
      Begin VB.TextBox txtSaldoAkhir 
         Height          =   495
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         Caption         =   "Saldo Akhir"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc adoDataOperasional 
      Height          =   495
      Left            =   6000
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
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
      RecordSource    =   "DataOperasional"
      Caption         =   "Data Operasional"
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
   Begin MSDataGridLib.DataGrid dtgDataOperasional 
      Bindings        =   "FrmDataOperasional.frx":0000
      Height          =   3375
      Left            =   240
      TabIndex        =   17
      Top             =   5520
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   5953
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
      ColumnCount     =   8
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
         DataField       =   "tanggal"
         Caption         =   "tanggal"
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
         DataField       =   "kode"
         Caption         =   "kode"
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
         DataField       =   "nama"
         Caption         =   "nama"
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
         DataField       =   "keterangan"
         Caption         =   "keterangan"
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
         DataField       =   "saldoAwal"
         Caption         =   "saldoAwal"
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
         DataField       =   "biayaKeluar"
         Caption         =   "biayaKeluar"
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
         DataField       =   "saldoAkhir"
         Caption         =   "saldoAkhir"
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
            ColumnWidth     =   1140,095
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
      EndProperty
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   2880
      TabIndex        =   16
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpTanggal 
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      Format          =   99418113
      CurrentDate     =   41519
   End
   Begin VB.TextBox txtNo 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtKode 
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtNama 
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox txtKeterangan 
      Height          =   1335
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3240
      Width           =   3855
   End
   Begin VB.TextBox txtBiayaKeluar 
      Height          =   495
      Left            =   7800
      TabIndex        =   7
      Top             =   960
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   1980
      Left            =   9120
      Picture         =   "FrmDataOperasional.frx":0021
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   4020
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Saldo Awal"
      Height          =   495
      Left            =   6120
      TabIndex        =   20
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "No"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "Tanggal"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "Kode"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Caption         =   "Nama"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   "Keterangan"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Biaya Keluar"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "FrmDataOperasional"
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
    dtpTanggal.Enabled = True
End Sub

Sub nonAktifTxt()
    For Each x In Me
        If TypeOf x Is TextBox Then
            x.Enabled = False
        End If
    Next
    dtpTanggal.Enabled = False
End Sub

Sub aktifBtn()
    btnHapus.Enabled = True
    btnSimpan.Enabled = True
    btnIsiSaldo.Enabled = True
End Sub

Sub nonAktifBtn()
    btnHapus.Enabled = False
    btnSimpan.Enabled = False
    btnIsiSaldo.Enabled = False
End Sub

Sub noAutomatis()
    On Error Resume Next
    Dim nomor As String
    adoDataOperasional.Recordset.Sort = "no"
    adoDataOperasional.RecordSource = "select * from DataOperasional"
    If adoDataOperasional.Recordset.RecordCount = 0 Then
        nomor = 1
    Else
        adoDataOperasional.Recordset.MoveLast
        nomor = adoDataOperasional.Recordset.Fields(0) + 1
    End If
    txtNo.Text = nomor
    On Error GoTo 0
End Sub

Sub simpan()
    On Error Resume Next
        With adoDataOperasional.Recordset
            .AddNew
            .Fields(0) = txtNo.Text
            .Fields(1) = dtpTanggal.Value
            .Fields(2) = txtKode.Text
            .Fields(3) = txtNama.Text
            .Fields(4) = txtKeterangan.Text
            .Fields(5) = txtJumlahSaldo.Text
            .Fields(6) = txtBiayaKeluar.Text
            .Fields(7) = txtSaldoAkhir.Text
            .Update
        End With
    On Error GoTo 0
End Sub

Sub isiSaldo()
    On Error Resume Next
        With adoDataOperasional.Recordset
            .AddNew
            .Fields(0) = txtNo.Text
            .Fields(1) = dtpTanggal.Value
            .Fields(2) = txtKode.Text
            .Fields(3) = txtNama.Text
            .Fields(4) = txtKeterangan.Text
            .Fields(5) = txtJumlahSaldo.Text
            If txtBiayaKeluar.Text = "" Then
                .Fields(6) = 0
            End If
            .Fields(7) = txtSaldoAkhir.Text
            .Update
        End With
    On Error GoTo 0
End Sub

Sub hapus()
    On Error Resume Next
    With adoDataOperasional.Recordset
        If .RecordCount = 0 Then
            MsgBox "Data sudah habis", , "Data Operasional"
        Else
            adoDataOperasional.Recordset.Delete
            adoDataOperasional.Recordset.Sort = "no"
        End If
    End With
    On Error GoTo 0
End Sub

Sub tampilSaldoAkhir()
    On Error Resume Next
    adoDataOperasional.RecordSource = "select saldoAkhir from DataOperasional"
    With adoDataOperasional.Recordset
        If .RecordCount = 0 Then
            txtSaldoAkhir.Text = 0
        Else
            .MoveLast
            txtSaldoAkhir.Text = .Fields(7)
        End If
    End With
    On Error Resume Next
End Sub

Private Sub btnBatal_Click()
    bersih
    noAutomatis
    tampilSaldoAkhir
    btnSimpan.Enabled = True
    btnUbah.Caption = "Ubah"
End Sub

Private Sub btnCetak_Click()
    bersih
    noAutomatis
End Sub

Private Sub btnHapus_Click()
    hapus
    bersih
    adoDataOperasional.Recordset.Sort = "no"
    noAutomatis
    tampilSaldoAkhir
End Sub

Private Sub btnHitungBk_Click()
    txtSaldoAkhir.Text = Val(txtSaldoAkhir.Text) - Val(txtBiayaKeluar.Text)
    txtJumlahSaldo.Text = 0
    btnHitungBk.Enabled = False
End Sub

Private Sub btnHitungSaldo_Click()
    txtSaldoAkhir.Text = Val(txtJumlahSaldo.Text) + Val(txtSaldoAkhir.Text)
    txtBiayaKeluar.Text = 0
    btnHitungSaldo.Enabled = False
End Sub

Private Sub btnSimpan_Click()
    simpan
    bersih
    txtKode.SetFocus
    noAutomatis
    tampilSaldoAkhir
End Sub

Private Sub btnUbah_Click()
    On Error Resume Next
    If btnUbah.Caption = "Ubah" Then
        btnUbah.Caption = "Perbarui"
        btnUbah.FontSize = 14
        btnUbah.FontBold = True
        btnSimpan.Enabled = False
        With adoDataOperasional.Recordset
            txtNo.Text = .Fields(0)
            dtpTanggal.Value = .Fields(1)
            txtKode.Text = .Fields(2)
            txtNama.Text = .Fields(3)
            txtKeterangan.Text = .Fields(4)
            txtBiayaKeluar.Text = .Fields(6)
            txtJumlahSaldo.Text = .Fields(5)
            .MovePrevious
            txtSaldoAkhir.Text = .Fields(7)
            .MoveNext
        End With
    Else
        btnUbah.Caption = "Ubah"
        btnUbah.FontBold = False
        btnUbah.FontSize = 8
        btnSimpan.Enabled = True
        With adoDataOperasional.Recordset
            .MoveLast
            .Delete
            .AddNew
            .Fields(0) = txtNo.Text
            .Fields(1) = dtpTanggal.Value
            .Fields(2) = txtKode.Text
            .Fields(3) = txtNama.Text
            .Fields(4) = txtKeterangan.Text
            .Fields(5) = txtJumlahSaldo.Text
            .Fields(6) = txtBiayaKeluar.Text
            .Fields(7) = txtSaldoAkhir.Text
            .Update
        End With
        bersih
        adoDataOperasional.Recordset.Sort = "no"
        noAutomatis
        tampilSaldoAkhir
    End If
    On Error GoTo 0
End Sub

Private Sub dtpTanggal_Click()
    txtKode.SetFocus
End Sub

Private Sub Form_Activate()
  noAutomatis
  tampilSaldoAkhir
  txtKode.SetFocus
End Sub

Private Sub txtBiayaKeluar_Click()
    btnHitungBk.Enabled = True
End Sub

Private Sub txtBiayaKeluar_GotFocus()
    btnHitungBk.Enabled = True
End Sub

Private Sub txtJumlahSaldo_Click()
    btnHitungSaldo.Enabled = True
End Sub

Private Sub txtJumlahSaldo_GotFocus()
    btnHitungSaldo.Enabled = True
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBiayaKeluar.SetFocus
    End If
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNama.SetFocus
    End If
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtKeterangan.SetFocus
    End If
End Sub

VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmDataKaryawan 
   BackColor       =   &H0000FF00&
   Caption         =   "Data Karyawan"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form2"
   ScaleHeight     =   6420
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNomorTelepon 
      Height          =   495
      Left            =   5640
      MaxLength       =   15
      TabIndex        =   4
      Top             =   120
      Width           =   2085
   End
   Begin MSDataGridLib.DataGrid dgDatakaryawan 
      Bindings        =   "FrmDataKaryawan.frx":0000
      Height          =   3255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
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
      ColumnCount     =   6
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
         DataField       =   "namaKaryawan"
         Caption         =   "namaKaryawan"
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
         DataField       =   "Divisi"
         Caption         =   "Divisi"
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
         DataField       =   "uangMakan"
         Caption         =   "uangMakan"
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
         DataField       =   "gajiPokok"
         Caption         =   "gajiPokok"
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
         DataField       =   "noTelepon"
         Caption         =   "noTelepon"
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
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1665,071
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtUangMakan 
      Height          =   495
      Left            =   5640
      MaxLength       =   15
      TabIndex        =   5
      Top             =   840
      Width           =   1485
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox cmbDivisi 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtGajiPokok 
      Height          =   495
      Left            =   5640
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1560
      Width           =   2325
   End
   Begin VB.TextBox txtNo 
      Height          =   495
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   1
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtNamaKaryawan 
      Height          =   495
      Left            =   1560
      MaxLength       =   22
      TabIndex        =   2
      Top             =   840
      Width           =   2300
   End
   Begin MSAdodcLib.Adodc adoDataKaryawan 
      Height          =   495
      Left            =   4200
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "DataKaryawan"
      Caption         =   "Data Karyawan"
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
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   7440
      Picture         =   "FrmDataKaryawan.frx":001E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2340
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "Nomor Telepon"
      Height          =   495
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Caption         =   "Uang Makan"
      Height          =   495
      Left            =   4200
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   "Gaji Pokok"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "No"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Nama Karyawan"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Divisi"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "FrmDataKaryawan"
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
    cmbDivisi.Text = ""
End Sub

Sub aktifTxt()
    For Each x In Me
        If TypeOf x Is TextBox Then
            x.Enabled = True
        End If
    Next
    cmbDivisi.Enabled = True
End Sub

Sub nonAktifTxt()
    For Each x In Me
        If TypeOf x Is TextBox Then
            x.Enabled = False
        End If
    Next
    cmbDivisi.Enabled = False
End Sub

Sub aktifBtn()
    btnHapus.Enabled = True
    btnSimpan.Enabled = True
End Sub

Sub nonAktifBtn()
    btnHapus.Enabled = False
    btnSimpan.Enabled = False
End Sub

Sub noAutomatis()
    On Error Resume Next
    Dim nomor As String
    adoDataKaryawan.Recordset.Sort = "no"
    adoDataKaryawan.RecordSource = "select * from dataKaryawan"
    If adoDataKaryawan.Recordset.RecordCount = 0 Then
        nomor = 1
    Else
        adoDataKaryawan.Recordset.MoveLast
        nomor = adoDataKaryawan.Recordset.Fields(0) + 1
    End If
    txtNo.Text = nomor
    On Error GoTo 0
End Sub
Sub isiDivisi()
    cmbDivisi.AddItem "Direktur Finance"
    cmbDivisi.AddItem "Akuntan"
    cmbDivisi.AddItem "Keuangan"
    cmbDivisi.AddItem "Koord Indosat"
    cmbDivisi.AddItem "Pm Non Jabo"
    cmbDivisi.AddItem "Koord Jabo"
    cmbDivisi.AddItem "Pajak & IT"
    cmbDivisi.AddItem "Control Document"
    cmbDivisi.AddItem "Pm Jabo"
    cmbDivisi.AddItem "Admin"
    cmbDivisi.AddItem "WareHouse"
    cmbDivisi.AddItem "Engineer"
    cmbDivisi.AddItem "Leader"
    cmbDivisi.AddItem "Teknisi+Driver"
    cmbDivisi.AddItem "Teknisi"
End Sub

Sub simpan()
    On Error Resume Next
    With adoDataKaryawan.Recordset
        .AddNew
        .Fields(0) = txtNo.Text
        .Fields(1) = txtNamaKaryawan.Text
        .Fields(2) = cmbDivisi.Text
        .Fields(3) = txtUangMakan.Text
        .Fields(4) = txtGajiPokok.Text
        .Fields(5) = txtNomorTelepon.Text
        .Update
    End With
    On Error GoTo 0
End Sub

Sub hapus()
    On Error Resume Next
    If adoDataKaryawan.Recordset.RecordCount = 0 Then
        MsgBox "Data sudah tidak ada", , "Hapus"
    Else
        If MsgBox("Ingin hapus data ini?", vbCritical + vbYesNo, "Peringatan") = vbYes Then
            adoDataKaryawan.Recordset.Delete
        Else
            Exit Sub
        End If
    End If
    On Error GoTo 0
End Sub
Private Sub btnHapus_Click()
    hapus
End Sub

Private Sub btnSimpan_Click()
    simpan
    bersih
    noAutomatis
    txtNamaKaryawan.SetFocus
End Sub

Private Sub btnTambah_Click()
    aktifBtn
    aktifTxt
    bersih
    isiDivisi
    noAutomatis
    txtNamaKaryawan.SetFocus
End Sub

Private Sub cmbDivisi_Click()
    If cmbDivisi.Text = "Dir Finance" Then
        txtGajiPokok.Text = 12000000
    ElseIf cmbDivisi.Text = "Akuntan" Then
        txtGajiPokok.Text = 1750000
    ElseIf cmbDivisi.Text = "Keuangan" Then
        txtGajiPokok.Text = 2500000
    ElseIf cmbDivisi.Text = "Koord Indosat" Then
        txtGajiPokok.Text = 1800000
    ElseIf cmbDivisi.Text = "Pm Non Jabo" Then
        txtGajiPokok.Text = 2500000
    ElseIf cmbDivisi.Text = "Koord Jabo" Then
        txtGajiPokok.Text = 1800000
    ElseIf cmbDivisi.Text = "Pajak & IT" Then
        txtGajiPokok.Text = 1800000
    ElseIf cmbDivisi.Text = "Control Document" Then
        txtGajiPokok.Text = 1400000
    ElseIf cmbDivisi.Text = "Pm Jabo" Then
        txtGajiPokok.Text = 1800000
    ElseIf cmbDivisi.Text = "Admin" Then
        txtGajiPokok.Text = 1300000
    ElseIf cmbDivisi.Text = "WareHouse" Then
        txtGajiPokok.Text = 1200000
    ElseIf cmbDivisi.Text = "Engineer" Then
        txtGajiPokok.Text = 1800000
    ElseIf cmbDivisi.Text = "Leader" Then
        txtGajiPokok.Text = 1500000
    ElseIf cmbDivisi.Text = "Teknisi+Driver" Then
        txtGajiPokok.Text = 1000000
    ElseIf cmbDivisi.Text = "Teknisi" Then
        txtGajiPokok.Text = 1300000
    Else
        txtGajiPokok.Text = 0
    End If
End Sub

Private Sub Form_Activate()
    nonAktifTxt
    nonAktifBtn
End Sub

VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmGajiKaryawan 
   BackColor       =   &H0000FF00&
   Caption         =   "Gaji Karyawan"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   13110
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnHitung 
      Caption         =   "Hitung"
      Height          =   495
      Left            =   10200
      TabIndex        =   20
      Top             =   960
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgGajiKaryawan 
      Bindings        =   "FrmGajiKaryawan.frx":0000
      Height          =   4095
      Left            =   240
      TabIndex        =   19
      Top             =   4080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7223
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
      BeginProperty Column03 
         DataField       =   "jumlahAlpha"
         Caption         =   "jumlahAlpha"
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
         DataField       =   "potonganLain"
         Caption         =   "potonganLain"
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
         DataField       =   "totalPotonganGaji"
         Caption         =   "totalPotonganGaji"
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
         DataField       =   "totalGaji"
         Caption         =   "totalGaji"
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
            ColumnWidth     =   915,024
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
   Begin MSComCtl2.DTPicker dtpTanggal 
      Height          =   495
      Left            =   1680
      TabIndex        =   18
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   99483649
      CurrentDate     =   41530
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Top             =   3360
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
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtJumlahAlpha 
      Height          =   495
      Left            =   7440
      TabIndex        =   12
      Top             =   240
      Width           =   915
   End
   Begin VB.TextBox txtPotonganLain 
      Height          =   495
      Left            =   7440
      TabIndex        =   10
      Top             =   960
      Width           =   2595
   End
   Begin VB.TextBox txtNo 
      Height          =   495
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin MSAdodcLib.Adodc adoDataKaryawan 
      Height          =   495
      Left            =   7920
      Top             =   5400
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "adoDataKaryawan"
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
   Begin MSAdodcLib.Adodc adoGajiKaryawan 
      Height          =   495
      Left            =   7920
      Top             =   4920
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      RecordSource    =   "GajiKaryawan"
      Caption         =   "adoGajiKaryawan"
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
   Begin VB.TextBox txtTotalGaji 
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   2400
      Width           =   2595
   End
   Begin VB.ComboBox cmbNamaKaryawan 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   2300
   End
   Begin VB.TextBox txtTotalPotonganGaji 
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   1680
      Width           =   2595
   End
   Begin VB.TextBox txtGajiPokok 
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   2000
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   9720
      Picture         =   "FrmGajiKaryawan.frx":001E
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "Tanggal"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   "Jumlah Alpha"
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Potongan Lain"
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "No"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "Total Gaji"
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Caption         =   "Total Potongan Gaji"
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Gaji Pokok"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Nama Karyawan"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "FrmGajiKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As New ADODB.Connection
Public rsGajiKaryawan As New ADODB.Recordset
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
    cmbNamaKaryawan.Enabled = True
End Sub

Sub nonAktifTxt()
    For Each x In Me
        If TypeOf x Is TextBox Then
            x.Enabled = False
        End If
    Next
    cmbNamaKaryawan.Enabled = False
End Sub

Sub aktifBtn()
    btnHapus.Enabled = True
    btnSimpan.Enabled = True
End Sub

Sub nonAktifBtn()
    btnHapus.Enabled = False
    btnSimpan.Enabled = False
End Sub

Sub koneksi()
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db.mdb;Persist Security Info=False"
    conn.Open
End Sub

Sub ambilData()
    On Error Resume Next
    koneksi

    Set rsGajiKaryawan = New ADODB.Recordset
    rsGajiKaryawan.Open "select namaKaryawan from dataKaryawan order by namaKaryawan ASC", conn

    Do While Not rsGajiKaryawan.EOF
        cmbNamaKaryawan.AddItem rsGajiKaryawan!namaKaryawan
        rsGajiKaryawan.MoveNext
    Loop
    rsGajiKaryawan.Close
    On Error GoTo 0
End Sub

Sub ambilGajiPokok()
    On Error Resume Next
    koneksi
    
    Set rsGajiKaryawan = New ADODB.Recordset
    rsGajiKaryawan.Open "select * from dataKaryawan where namaKaryawan = '" + cmbNamaKaryawan.Text + "'", conn
    With rsGajiKaryawan
        txtGajiPokok.Text = rsGajiKaryawan!gajiPokok
    End With
    On Error GoTo 0
End Sub

Sub ambilAlpha()
    On Error Resume Next
    koneksi

    Set rsGajiKaryawan = New ADODB.Recordset
    rsGajiKaryawan.Open "select count(absenAlpha) as jumlahAlpha from DataAbsensi where namaKaryawan = '" + cmbNamaKaryawan.Text + "' and YEAR(tanggal)='" & dtpTanggal.Year & "' and MONTH(tanggal)='" & dtpTanggal.Month & "' and absenAlpha='Alpha'", conn
    With rsGajiKaryawan
        txtJumlahAlpha.Text = rsGajiKaryawan!jumlahAlpha
    End With
    On Error GoTo 0
End Sub

Sub noAutomatis()
    On Error Resume Next
    Dim nomor As String
    adoGajiKaryawan.RecordSource = "select * from gajiKaryawan"
    If adoGajiKaryawan.Recordset.RecordCount = 0 Then
        nomor = 1
    Else
        adoGajiKaryawan.Recordset.MoveLast
        nomor = adoGajiKaryawan.Recordset.Fields(0) + 1
    End If
    txtNo.Text = nomor
    On Error GoTo 0
End Sub

Sub simpan()
    On Error Resume Next
    With adoGajiKaryawan.Recordset
        .AddNew
        .Fields(0) = txtNo.Text
        .Fields(1) = cmbNamaKaryawan.Text
        .Fields(2) = txtGajiPokok.Text
        .Fields(3) = txtJumlahAlpha
        .Fields(4) = txtPotonganLain.Text
        .Fields(5) = txtTotalPotonganGaji.Text
        .Fields(6) = txtTotalGaji.Text
        .Fields(7) = dtpTanggal.Value
        .Update
    End With
    On Error GoTo 0
End Sub
Sub hapus()
    On Error Resume Next
    If adoGajiKaryawan.Recordset.RecordCount = 0 Then
        MsgBox "Data sudah tidak ada", , "Hapus"
    Else
        If MsgBox("Ingin hapus data ini?", vbCritical + vbYesNo, "Peringatan") = vbYes Then
            adoGajiKaryawan.Recordset.Delete
        Else
            Exit Sub
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub btnBatal_Click()
    bersih
    cmbNamaKaryawan.Clear
    noAutomatis
End Sub

Private Sub btnHapus_Click()
    hapus
    bersih
    adoGajiKaryawan.Recordset.Sort = "no"
    noAutomatis
End Sub

Private Sub btnHitung_Click()
    If txtPotonganLain.Text = "" Then
        txtPotonganLain.Text = 0
        txtTotalPotonganGaji.Text = ((Val(txtGajiPokok.Text) / 20) * Val(txtJumlahAlpha.Text)) + Val(txtPotonganLain.Text)
        txtTotalGaji.Text = Val(txtGajiPokok.Text) - ((Val(txtGajiPokok.Text) / 20) * Val(txtJumlahAlpha.Text)) - Val(txtPotonganLain.Text)
    Else
        txtTotalPotonganGaji.Text = ((Val(txtGajiPokok.Text) / 20) * Val(txtJumlahAlpha.Text)) + Val(txtPotonganLain.Text)
        txtTotalGaji.Text = Val(txtGajiPokok.Text) - ((Val(txtGajiPokok.Text) / 20) * Val(txtJumlahAlpha.Text)) - Val(txtPotonganLain.Text)
    End If
End Sub

Private Sub btnSimpan_Click()
    simpan
    bersih
    noAutomatis
    cmbNamaKaryawan.Clear
    cmbNamaKaryawan.SetFocus
End Sub

Private Sub btnTambah_Click()
    On Error Resume Next
    bersih
    cmbNamaKaryawan.Clear
    aktifBtn
    aktifTxt
    noAutomatis
    ambilData
    On Error GoTo 0
End Sub

Private Sub cmbNamaKaryawan_Click()
    ambilGajiPokok
    ambilAlpha
End Sub

Private Sub Form_Activate()
    noAutomatis
    ambilData
    dtpTanggal.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    conn.Close
    On Error GoTo 0
End Sub

Private Sub optGaji_Click()
    adoGajiKaryawan.Recordset.Sort = "totalGaji"
End Sub

Private Sub optNama_Click()
    adoGajiKaryawan.Recordset.Sort = "namaKaryawan"
End Sub

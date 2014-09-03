VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmDataAbsensi 
   BackColor       =   &H0000FF00&
   Caption         =   "Data Absensi"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   8475
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid dgDataAbsensi 
      Bindings        =   "FrmDataAbsensi.frx":0000
      Height          =   4455
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7858
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
         DataField       =   "absenHadir"
         Caption         =   "absenHadir"
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
         DataField       =   "absenAlpha"
         Caption         =   "absenAlpha"
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
      BeginProperty Column05 
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
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoDataAbsensi 
      Height          =   495
      Left            =   5040
      Top             =   3720
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
      RecordSource    =   "DataAbsensi"
      Caption         =   "Data Absensi"
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
   Begin VB.CommandButton btnBatal 
      Caption         =   "Batal"
      Height          =   495
      Left            =   1440
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtUangMakan 
      Height          =   495
      Left            =   6360
      TabIndex        =   12
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtNo 
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   240
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpTanggal 
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   99221505
      CurrentDate     =   41521
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton optAlpha 
      BackColor       =   &H0000FF00&
      Caption         =   "Alpha"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton optHadir 
      BackColor       =   &H0000FF00&
      Caption         =   "Hadir"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cmbNamaKaryawan 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   5040
      Picture         =   "FrmDataAbsensi.frx":001D
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3180
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "Uang Makan"
      Height          =   495
      Left            =   4920
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   "No"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Tanggal"
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Absen"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Nama Karyawan"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "FrmDataAbsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rsDataAbsensi As New ADODB.Recordset
Dim qry As String

Sub bersih()
    cmbNamaKaryawan.Text = ""
    txtUangMakan.Text = 0
    optAlpha.Value = False
    optHadir.Value = False
End Sub

Sub koneksi()
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db.mdb;Persist Security Info=False"
    conn.Open
End Sub

Sub ambilData()
    On Error Resume Next
    koneksi
    rsDataAbsensi.Open "select namaKaryawan from dataKaryawan order by namaKaryawan ASC", conn

    Do While Not rsDataAbsensi.EOF
        cmbNamaKaryawan.AddItem rsDataAbsensi!namaKaryawan
        rsDataAbsensi.MoveNext
    Loop
    rsDataAbsensi.Close
    On Error GoTo 0
End Sub

Sub ambilDataUangMakan()
    On Error Resume Next
    koneksi
    rsDataAbsensi.Open "select uangMakan from dataKaryawan where namaKaryawan = '" + cmbNamaKaryawan.Text + "'", conn
    txtUangMakan.Text = rsDataAbsensi!uangMakan
    rsDataAbsensi.Close
    On Error GoTo 0
End Sub
Sub noAutomatis()
    On Error Resume Next
    Dim nomor As String
    adoDataAbsensi.Recordset.Sort = "no"
    adoDataAbsensi.RecordSource = "select * from DataAbsensi"
    If adoDataAbsensi.Recordset.RecordCount = 0 Then
        nomor = 1
    Else
        adoDataAbsensi.Recordset.MoveLast
        nomor = adoDataAbsensi.Recordset.Fields(0) + 1
    End If
    txtNo.Text = nomor
    On Error GoTo 0
End Sub

Sub simpan()
    On Error Resume Next
    With adoDataAbsensi.Recordset
        .AddNew
        .Fields(0) = txtNo.Text
        .Fields(1) = cmbNamaKaryawan.Text
        If optHadir.Value = True Then
            .Fields(2) = "Hadir"
            .Fields(3) = ""
        ElseIf optAlpha.Value = True Then
            .Fields(2) = ""
            .Fields(3) = "Alpha"
        End If
        .Fields(4) = txtUangMakan.Text
        .Fields(5) = dtpTanggal.Value
        .Update
    End With
    On Error GoTo 0
End Sub

Sub hapus()
    On Error Resume Next
    If adoDataAbsensi.Recordset.RecordCount = 0 Then
        MsgBox "Data habis!", , "Hapus Data"
    Else
        If MsgBox("Ingin hapus data ini?", vbCritical + vbYesNo, "Hapus data") = vbYes Then
            adoDataAbsensi.Recordset.Delete
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub btnBatal_Click()
    bersih
End Sub

Private Sub btnHapus_Click()
    hapus
    bersih
    noAutomatis
End Sub

Private Sub btnSimpan_Click()
    simpan
    bersih
    noAutomatis
End Sub

Private Sub Form_Load()
    bersih
    ambilData
    noAutomatis
End Sub

Private Sub optAlpha_Click()
    txtUangMakan.Text = 0
End Sub

Private Sub optHadir_Click()
    ambilDataUangMakan
End Sub


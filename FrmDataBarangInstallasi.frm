VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmDataBarangInstallasi 
   BackColor       =   &H0000FF00&
   Caption         =   "Data Barang Installasi"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15840
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   15840
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11160
      Top             =   720
   End
   Begin VB.TextBox txtPurchaseDoc 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox txtNoGac 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker dtpReceivedDate 
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   99549185
      CurrentDate     =   41522
   End
   Begin MSAdodcLib.Adodc adoDataBarang 
      Height          =   495
      Left            =   9480
      Top             =   6720
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "DataBarangInstallasi"
      Caption         =   "Data Barang"
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
   Begin MSDataGridLib.DataGrid dtgDataBarang 
      Bindings        =   "FrmDataBarangInstallasi.frx":0000
      Height          =   2295
      Left            =   120
      TabIndex        =   31
      Top             =   6000
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   4048
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
      ColumnCount     =   14
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
         DataField       =   "noGAC"
         Caption         =   "noGAC"
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
         DataField       =   "linkNameSiteID"
         Caption         =   "linkNameSiteID"
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
         DataField       =   "requestor"
         Caption         =   "requestor"
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
         DataField       =   "receivedDate"
         Caption         =   "receivedDate"
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
         DataField       =   "purchaseDoc"
         Caption         =   "purchaseDoc"
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
         DataField       =   "item"
         Caption         =   "item"
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
         DataField       =   "sapCode"
         Caption         =   "sapCode"
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
         DataField       =   "shortText"
         Caption         =   "shortText"
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
         DataField       =   "qty"
         Caption         =   "qty"
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
         DataField       =   "totalDeliveredQty"
         Caption         =   "totalDeliveredQty"
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
         DataField       =   "serialNo"
         Caption         =   "serialNo"
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
         DataField       =   "ewo"
         Caption         =   "ewo"
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
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1140,095
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbShortText 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7920
      TabIndex        =   9
      Top             =   1680
      Width           =   5535
   End
   Begin VB.TextBox txtLinkNamesiteId 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   1600
   End
   Begin VB.TextBox txtNamaProject 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   3600
   End
   Begin VB.TextBox txtNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1100
   End
   Begin VB.TextBox txtItem 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7920
      TabIndex        =   7
      Top             =   240
      Width           =   1100
   End
   Begin VB.TextBox txtSapCode 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7920
      TabIndex        =   8
      Top             =   960
      Width           =   1600
   End
   Begin VB.TextBox txtRequestor 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   3600
   End
   Begin VB.TextBox txtQuantity 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtTotalDelivered 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7920
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtSerialNo 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7920
      TabIndex        =   12
      Top             =   3840
      Width           =   1500
   End
   Begin VB.TextBox txtEwo 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7920
      TabIndex        =   13
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   2760
      TabIndex        =   16
      Top             =   5220
      Width           =   1215
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   5220
      Width           =   1215
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   5220
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   10920
      Picture         =   "FrmDataBarangInstallasi.frx":001C
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblTanggal 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      TabIndex        =   33
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblJam 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   32
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "SAP Code"
      Height          =   495
      Left            =   6480
      TabIndex        =   30
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Item"
      Height          =   495
      Left            =   6480
      TabIndex        =   29
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   "Purchase Doc"
      Height          =   495
      Left            =   240
      TabIndex        =   28
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Caption         =   "Received Date"
      Height          =   495
      Left            =   240
      TabIndex        =   27
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "Requestor"
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "Link Namesite ID"
      Height          =   495
      Left            =   240
      TabIndex        =   25
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "Nama Project"
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FF00&
      Caption         =   "No GAC"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FF00&
      Caption         =   "No"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000FF00&
      Caption         =   "Quantity"
      Height          =   495
      Left            =   6480
      TabIndex        =   21
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H0000FF00&
      Caption         =   "Short Text"
      Height          =   495
      Left            =   6480
      TabIndex        =   20
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Total Delivered Quantity"
      Height          =   495
      Left            =   6480
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000FF00&
      Caption         =   "Serial No"
      Height          =   495
      Left            =   6480
      TabIndex        =   18
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000FF00&
      Caption         =   "Ewo"
      Height          =   495
      Left            =   6480
      TabIndex        =   17
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "FrmDataBarangInstallasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rsDataBarang As New ADODB.Recordset
Dim qry As String

Sub bersih()
    txtItem.Text = ""
    txtSapCode.Text = ""
    cmbShortText.Text = ""
    txtQuantity.Text = ""
    txtTotalDelivered.Text = ""
    txtSerialNo.Text = ""
End Sub

Sub bersihTambah()
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
    dtpReceivedDate.Enabled = True
    cmbShortText.Enabled = True
End Sub

Sub nonAktifTxt()
    For Each x In Me
        If TypeOf x Is TextBox Then
            x.Enabled = False
        End If
    Next
    dtpReceivedDate.Enabled = False
    cmbShortText.Enabled = False
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
End Sub

Sub simpan()
    On Error Resume Next
    With adoDataBarang.Recordset
        .AddNew
        .Fields(0) = txtNo.Text
        .Fields(1) = txtNoGac.Text
        .Fields(2) = txtNamaProject.Text
        .Fields(3) = txtLinkNamesiteId.Text
        .Fields(4) = txtRequestor.Text
        .Fields(5) = dtpReceivedDate.Value
        .Fields(6) = txtPurchaseDoc.Text
        .Fields(7) = txtItem.Text
        .Fields(8) = txtSapCode.Text
        .Fields(9) = cmbShortText.Text
        .Fields(10) = txtQuantity.Text
        .Fields(11) = txtTotalDelivered.Text
        .Fields(12) = txtSerialNo.Text
        .Fields(13) = txtEwo.Text
        .Update
    End With
    On Error GoTo 0
End Sub

Sub koneksi()
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db.mdb;Persist Security Info=False"
    conn.Open
End Sub

Sub ambilDataShortText()
    On Error Resume Next
    koneksi

    Set rsDataBarang = New ADODB.Recordset
    rsDataBarang.Open "select shortText from DataBarangInstallasi group by shortText having count(*) >= 1", conn

    Do While Not rsDataBarang.EOF
        cmbShortText.AddItem rsDataBarang!shortText
        rsDataBarang.MoveNext
    Loop
    rsDataBarang.Close
    On Error GoTo 0
End Sub

Sub hapus()
    On Error Resume Next
    If adoDataBarang.Recordset.RecordCount = 0 Then
        MsgBox "Data sudah tidak ada", , "Hapus"
    Else
        If MsgBox("Ingin hapus data ini?", vbCritical + vbYesNo, "Peringatan") = vbYes Then
            adoDataBarang.Recordset.Delete
        Else
            Exit Sub
        End If
    End If
    On Error GoTo 0
End Sub

Sub noAutomatis()
    On Error Resume Next
    Dim nomor As String
    adoDataBarang.Recordset.Sort = "no"
    adoDataBarang.RecordSource = "select * from DataBarangInstallasi"
    If adoDataBarang.Recordset.RecordCount = 0 Then
        nomor = 1
    Else
        adoDataBarang.Recordset.MoveLast
        nomor = adoDataBarang.Recordset.Fields(0) + 1
    End If
    txtNo.Text = nomor
    On Error GoTo 0
End Sub

Private Sub btnHapus_Click()
    On Error Resume Next
    hapus
    adoDataBarang.Recordset.Sort = "no"
    noAutomatis
    On Error GoTo 0
End Sub

Private Sub btnSimpan_Click()
    simpan
    bersih
    noAutomatis
    ambilDataShortText
    txtItem.SetFocus
End Sub

Private Sub btnTambah_Click()
    On Error Resume Next
    aktifBtn
    aktifTxt
    bersihTambah
    noAutomatis
    txtNoGac.SetFocus
    ambilDataShortText
    On Error GoTo 0
End Sub

Private Sub cmbShortText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtQuantity.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    nonAktifTxt
    nonAktifBtn
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    conn.Close
    On Error GoTo 0
End Sub

Private Sub optSapCode_Click()
    On Error Resume Next
        adoDataBarang.Recordset.Sort = "sapCOde"
    On Error GoTo 0
End Sub

Private Sub optTanggal_Click()
    On Error Resume Next
        adoDataBarang.Recordset.Sort = "receivedDate"
    On Error GoTo 0
End Sub

Private Sub Timer1_Timer()
    lblJam.Caption = Time
    lblTanggal.Caption = Date
End Sub

Private Sub txtEwo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnSimpan.SetFocus
    End If
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSapCode.SetFocus
    End If
End Sub

Private Sub txtLinkNamesiteId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRequestor.SetFocus
    End If
End Sub

Private Sub txtNamaProject_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtLinkNamesiteId.SetFocus
    End If
End Sub

Private Sub txtNoGac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNamaProject.SetFocus
    End If
End Sub

Private Sub txtPurchaseDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtItem.SetFocus
    End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTotalDelivered.Text = txtQuantity.Text
        txtSerialNo.SetFocus
    End If
End Sub

Private Sub txtRequestor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPurchaseDoc.SetFocus
    End If
End Sub

Private Sub txtSapCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbShortText.SetFocus
    End If
End Sub

Private Sub txtSerialNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEwo.SetFocus
    End If
End Sub

Private Sub txtTotalDelivered_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSerialNo.SetFocus
    End If
End Sub

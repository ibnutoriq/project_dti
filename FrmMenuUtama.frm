VERSION 5.00
Begin VB.Form FrmMenuUtama 
   BackColor       =   &H0000FF00&
   Caption         =   "Menu Utama"
   ClientHeight    =   5445
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   7260
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form3"
   Picture         =   "FrmMenuUtama.frx":0000
   ScaleHeight     =   5445
   ScaleWidth      =   7260
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   10695
      Left            =   -120
      Picture         =   "FrmMenuUtama.frx":3937
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20220
   End
   Begin VB.Menu mnMaster 
      Caption         =   "Master"
      Begin VB.Menu smnDataKaryawan 
         Caption         =   "Data Karyawan"
         Shortcut        =   {F1}
      End
      Begin VB.Menu smnGajiKaryawan 
         Caption         =   "Gaji Karyawan"
         Shortcut        =   {F2}
      End
      Begin VB.Menu smnPurchaseOrder 
         Caption         =   "Purchase Order"
         Shortcut        =   {F3}
      End
      Begin VB.Menu smnDataBarangInstallasi 
         Caption         =   "Data Barang Installasi"
         Shortcut        =   {F4}
      End
      Begin VB.Menu smnDataOperational 
         Caption         =   "Data Operational"
         Shortcut        =   {F5}
      End
      Begin VB.Menu smnDataAbsensi 
         Caption         =   "Data Absensi"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "FrmMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
    Image1.Width = Width
    Image1.Height = Height
End Sub

Private Sub mnAbout_Click()
    FrmAbout.Show
End Sub

Private Sub mnExit_Click()
    If MsgBox("Ingin Keluar?", vbQuestion + vbYesNo, "Keluar") = vbYes Then
        End
    End If
End Sub

Private Sub smnDataAbsensi_Click()
    FrmDataAbsensi.Show
End Sub

Private Sub smnDataBarangInstallasi_Click()
    FrmDataBarangInstallasi.Show
End Sub

Private Sub smnDataKaryawan_Click()
    FrmDataKaryawan.Show
End Sub

Private Sub smnDataOperational_Click()
    FrmDataOperasional.Show
End Sub

Private Sub smnGajiKaryawan_Click()
    FrmGajiKaryawan.Show
End Sub

Private Sub smnPurchaseOrder_Click()
    FrmPurchaseOrder.Show
End Sub

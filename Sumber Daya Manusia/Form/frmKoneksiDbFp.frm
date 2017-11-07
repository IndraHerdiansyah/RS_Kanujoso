VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmKoneksiDb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Setting Database Finger Print"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   Icon            =   "frmKoneksiDbFp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   5415
      Begin VB.CommandButton cmdIntegrasidata 
         Caption         =   "Integrasi Data"
         Height          =   495
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpDari 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   117637121
         CurrentDate     =   41835
      End
      Begin MSComCtl2.DTPicker dtpSampai 
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   117309441
         CurrentDate     =   41835
      End
      Begin VB.Label Label6 
         Caption         =   "Dari tanggal"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Sampai tanggal"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSaveReg 
      Caption         =   "&Test Koneksi"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5415
      Begin VB.TextBox txtDatabaseName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtServerName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "IP. Addres"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   315
         Width           =   735
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3840
      Picture         =   "frmKoneksiDbFp.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKoneksiDbFp.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmKoneksiDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private context As Core_FingerPrint.DeviceManager
Public guidId As String
Private fso As New Scripting.FileSystemObject

Private Sub subLoadDBFP()
    txtServerName.Text = funcGetFromINI("Database Finger Print", "Nama Server", "", strFileSettingDBFP)
    txtDatabaseName.Text = funcGetFromINI("Database Finger Print", "Nama Database", "", strFileSettingDBFP)
End Sub

Private Sub subDefaultText(ByVal TextBoxObj As TextBox, ByVal DefaultText As String)
    With TextBoxObj
        .ForeColor = vbGrayText
        .Text = DefaultText
    End With
End Sub

Private Sub cmdBatal_Click()
    TglAkhir = dtpSampai.Value
    TglAwal = dtpDari.Value
    Unload Me
End Sub

Private Sub cmdIntegrasidata_Click()

    If Periksa("text", txtServerName, "Nama pegawai kosong") = False Then Exit Sub
    If Periksa("text", txtDatabaseName, "Jenis kelamin kosong") = False Then Exit Sub
    subDeletemAttendance '//yayang.agus 2014-08-22
'    context.SetConnectionString (Replace(Replace(dbConn.ConnectionString, "Provider=SQLNCLI10.1;", ""), "Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=" & UCase(strNamaHostLocal) & ";Use Encryption for Data=False;Tag with column collation when possible=False;MARS Connection=False;DataTypeCompatibility=80;Trust Server Certificate=False", ""))
'    guidId = context.GetData(Format(dtpDari.Value, "dd/MM/yyyy"), Format(dtpSampai.Value, "dd/MM/yyyy"))
    TglAkhir = dtpSampai.Value
    TglAwal = dtpDari.Value
    Unload Me
    'MsgBox "Data dari Tgl" + dtpDari.Value + "Sampai Tgl" + dtpSampai.Value + "Berhasil diambil", vbInformation
End Sub

Private Sub subDeletemAttendance()
    strSQL = "delete from mAttendance "
    Call msubRecFO(rs, strSQL)
End Sub


Private Sub cmdSaveReg_Click()
Dim data As Boolean


'Set context = New Core_FingerPrint.DeviceManager
data = context.connection(txtServerName.Text, CInt(txtDatabaseName.Text))
If (data = True) Then
    Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Ip Finger Print ", txtServerName.Text)
    Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Port Finger Print ", txtDatabaseName.Text)
    MsgBox "Data berhasil terkoneksi"
End If
'MsgBox data
cmdIntegrasidata.Enabled = True
dtpDari.SetFocus





'    On Error Resume Next
'    Dim dbConnFP As New ADODB.connection
'    Dim myConSTR As String
'    Screen.MousePointer = vbHourglass
'    dbConnFP.CursorLocation = adUseServer
'    myConSTR = "Data Source=" & txtDatabaseName.Text & ";Data Source=" & txtServerName.Text
'    'myConSTR = "Provider=SQLOLEDB10;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & txtDatabaseName.Text & ";Data Source=" & txtServerName.Text
'    dbConnFP.Open myConSTR
'    Screen.MousePointer = vbDefault
'    If Err Then
'        MsgBox "SQL Connection Failed: " & Err.Description, vbCritical, "Error.."
'        Set dbConnFP = Nothing
'        txtServerName.SetFocus
'
'    Else
'        Dim pesan As VbMsgBoxResult
'
'        Call funcAddToINI("Database Finger Print", "Nama Server", txtServerName.Text, strFileSettingDBFP)
'        Call funcAddToINI("Database Finger Print", "Nama Database", txtDatabaseName.Text, strFileSettingDBFP)
'        Set dbConnFP = Nothing
'        pesan = MsgBox("Koneksi database berhasil!" & vbCrLf & "Tutup form Setting Database Finger Print?" _
'        , vbInformation Or vbYesNo, "Success..")
'        If pesan = vbYes Then
'            Unload Me
'        End If
'
'    End If
End Sub

Private Sub dtpDari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then dtpSampai.SetFocus
TglAkhir = dtpSampai.Value
TglAwal = dtpDari.Value
End Sub


Private Sub dtpSampai_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdIntegrasidata.SetFocus
TglAkhir = dtpSampai.Value
TglAwal = dtpDari.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    cmdIntegrasidata.Enabled = False
    dtpDari = Now
    dtpSampai = Now
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    txtServerName.Text = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Ip Finger Print ")
    txtDatabaseName.Text = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Port Finger Print ")
    'If fso.FileExists(strFileSettingDBFP) Then
    '    Call subLoadDBFP
    'Else
    '    Call subDefaultText(Me.txtDatabaseName, "[Tidak ada nama database]")
    '    Call subDefaultText(Me.txtServerName, "[Tidak ada nama server]")
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With frmAbsensiPegawai_OffLine
        .tmrAbsensi.Enabled = True
        .subLoadDBFP
    End With
End Sub

Private Sub txtDatabaseName_GotFocus()
    If txtDatabaseName.Text = "[Tidak ada nama database]" Then
        txtDatabaseName.Text = ""
        txtDatabaseName.ForeColor = vbBlack
    End If
End Sub

Private Sub txtDatabaseName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSaveReg.SetFocus
End Sub

Private Sub txtDatabaseName_LostFocus()
    If Trim(txtDatabaseName.Text) = "" Then
        Call subDefaultText(Me.txtDatabaseName, "[Tidak ada nama database]")
    Else
        txtDatabaseName.Text = UCase(txtDatabaseName.Text)
    End If
End Sub

Private Sub txtServerName_GotFocus()
    If Me.txtServerName.Text = "[Tidak ada nama server]" Then
        Me.txtServerName.Text = ""
        Me.txtServerName.ForeColor = vbBlack
    End If
End Sub

Private Sub txtServerName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtServerName.SetFocus
End Sub

Private Sub txtServerName_LostFocus()
    If Trim(txtServerName.Text) = "" Then
        Call subDefaultText(txtServerName, "[Tidak ada nama server]")
    Else
        txtServerName.Text = UCase(txtServerName.Text)
    End If
End Sub

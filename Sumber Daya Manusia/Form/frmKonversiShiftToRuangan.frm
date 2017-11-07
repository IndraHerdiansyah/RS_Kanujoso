VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Begin VB.Form frmKonversiShiftToRuangan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medifirst 2000 - Konversi Shift Kerja To Ruangan"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKonversiShiftToRuangan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
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
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmKonversiShiftToRuangan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6000
      Picture         =   "frmKonversiShiftToRuangan.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKonversiShiftToRuangan.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmKonversiShiftToRuangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilterPegawai As String
Dim strFilterPasien As String
Dim intJmlPegawai As Integer
Dim intJmlPAsien As Integer
Dim subIdPegawai As String
Dim subNoCM As String

Private Sub cmdBatal_Click()
    Call clearData
    Call loadDcSource
    Call loadGridSource
End Sub

Private Sub cmdHapus_Click()
On Error GoTo hell
If txtNoUrut.Text = "" Then MsgBox "Silahkan pilih data yang akan dihapus", vbCritical, "Konfirmasi"
Set adoComm = New ADODB.Command
    With adoComm
        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AUD_ConvertPegawaiToPasien"
        
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, subIdPegawai)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, subNoCM)
        .Parameters.Append .CreateParameter("Hubungan", adChar, adParamInput, 2, dcShift.BoundText)
        .Parameters.Append .CreateParameter("NoUrut", adTinyInt, adParamInput, , txtNoUrut.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "D")
        
        .Execute
        MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
        cmdBatal_Click
        
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errSave
    
    If Periksa("text", txtPegawai, "Silahkan isi No. ID / Nama Pegawai") = False Then Exit Sub
    If Periksa("text", txtRuangan, "Silahkan isi No.CM / Nama Pasien") = False Then Exit Sub
    If Periksa("datacombo", dcShift, "Silahkan pilih Hubungan Keluarga") = False Then Exit Sub
    
    Set adoComm = New ADODB.Command
    With adoComm
        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AUD_ConvertPegawaiToPasien"
        
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, subIdPegawai)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, subNoCM)
        .Parameters.Append .CreateParameter("Hubungan", adChar, adParamInput, 2, dcShift.BoundText)
        If txtNoUrut.Text = "" Then
        .Parameters.Append .CreateParameter("NoUrut", adTinyInt, adParamInput, , Null)
        Else
        .Parameters.Append .CreateParameter("NoUrut", adTinyInt, adParamInput, , txtNoUrut.Text)
        End If
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        
        .Execute
        
        MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
        cmdBatal_Click
        
    End With
    Exit Sub
    
errSave:
    Call msubPesanError(" cmdSimpan_Click")
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcRekeningImpact_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then optDebetImpact.SetFocus
End Sub

Private Sub dcRekening_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcRekeningImpact.SetFocus
End Sub

Private Sub dcShift_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dgKonversi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    With dgKonversi
        txtNoUrut.Text = .Columns(0).Value
        subIdPegawai = .Columns(1).Value
        txtPegawai.Text = .Columns(2).Value
        subNoCM = .Columns(3).Value
        txtRuangan.Text = .Columns(4).Value
        dcShift.BoundText = .Columns(5).Value
    End With
    fraPegawai.Visible = False
    fraRuangan.Visible = False
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
'    Call openConnection
    Call loadDcSource
    Call loadGridSource
End Sub

Private Sub loadDcSource()
On Error GoTo errLoad
    Call msubDcSource(dcShift, rs, "SELECT KdShift, NamaShift FROM ShiftKerja order by NamaShift")
Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Public Sub loadGridSource()
On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "SELECT * FROM V_KonversiPegawaiKePasien"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgKonversi.DataSource = rs
    Call subSetGrid
Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub subSetGrid()
    With dgKonversi
        .Columns(0).Width = 1000
        .Columns(1).Width = 0
        .Columns(2).Width = 3000
        .Columns(3).Width = 0
        .Columns(4).Width = 3000
        .Columns(5).Width = 0
        .Columns(6).Width = 4000
    End With
End Sub

Public Sub clearData()
    txtNoUrut.Text = ""
    txtPegawai.Text = ""
    txtRuangan.Text = ""
    dcShift.BoundText = ""
    txtPegawai.SetFocus
    fraPegawai.Visible = False
    fraRuangan.Visible = False
End Sub

Private Sub txtPegawai_Change()
On Error GoTo errLoad
    If subTampil = True Then Exit Sub
    strFilterPegawai = "WHERE NamaLengkap like '%" & txtPegawai.Text & "%' or IdPegawai like '%" & txtPegawai.Text & "%'"
    fraPegawai.Visible = True
    Call subLoadPegawai
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadPegawai()
On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select top 100 IdPegawai, NamaLengkap from DataPegawai " & strFilterPegawai
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlPegawai = rs.RecordCount
    Set dgPegawai.DataSource = rs
    With dgPegawai
        .Columns(0).Caption = "Id Pegawai"
        .Columns(0).Width = 1200
        .Columns(1).Caption = "Nama Lengkap"
        .Columns(1).Width = 3000
    End With
    fraPegawai.Left = 1320
    fraPegawai.Top = 960
Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub txtPegawai_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        If intJmlPegawai = 0 Then Exit Sub
        If fraPegawai.Visible = True Then
            dgPegawai.SetFocus
        Else
            dcShift.SetFocus
            Set rs = Nothing
            strSQL = "select top 100 IdPegawai from DataPegawai where NamaLengkap='" & txtPegawai.Text & "'" & strFilterPegawai
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            subIdPegawai = rs.Fields(0).Value
        End If
    End If
    If KeyAscii = 27 Then
        fraPegawai.Visible = False
    End If
Exit Sub
hell:
End Sub

Private Sub dgPegawai_DblClick()
    Call dgPegawai_KeyPress(13)
End Sub

Private Sub dgPegawai_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        txtPegawai.Text = dgPegawai.Columns(1).Value
        subIdPegawai = dgPegawai.Columns(0).Value
        fraPegawai.Visible = False
        txtRuangan.SetFocus
    End If
    If KeyAscii = 27 Then
        fraPegawai.Visible = False
    End If
Exit Sub
errLoad:
End Sub

Private Sub txtRuangan_Change()
On Error GoTo errLoad
    If subTampil = True Then Exit Sub
    strFilterPasien = "WHERE NamaLengkap like '%" & txtRuangan.Text & "%' or NoCm like '%" & txtRuangan.Text & "%'"
    fraRuangan.Visible = True
    Call subLoadpasien
Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub subLoadpasien()
On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select top 50 NoCm, NamaLengkap from Pasien " & strFilterPasien
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlPAsien = rs.RecordCount
    Set dgRuangan.DataSource = rs
    With dgRuangan
        .Columns(0).Caption = "No. CM"
        .Columns(0).Width = 1200
        .Columns(1).Caption = "Nama Pasien"
        .Columns(1).Width = 3000
    End With
    fraRuangan.Left = 3960
    fraRuangan.Top = 960
Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub txtRuangan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        If intJmlPAsien = 0 Then Exit Sub
        If fraRuangan.Visible = True Then
            dgRuangan.SetFocus
        Else
            dcShift.SetFocus
            Set rs = Nothing
            strSQL = "select top 50 NoCm from Pasien where NamaLengkap='" & txtRuangan.Text & "'" & strFilterPasien
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            subNoCM = rs.Fields(0).Value
        End If
    End If
    If KeyAscii = 27 Then
        fraRuangan.Visible = False
    End If
Exit Sub
hell:
End Sub

Private Sub dgpasien_DblClick()
    Call dgpasien_KeyPress(13)
End Sub

Private Sub dgpasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRuangan.Text = dgRuangan.Columns(1).Value
        subNoCM = dgRuangan.Columns(0).Value
        fraRuangan.Visible = False
        dcShift.SetFocus
    End If
    If KeyAscii = 27 Then
        fraRuangan.Visible = False
    End If
End Sub

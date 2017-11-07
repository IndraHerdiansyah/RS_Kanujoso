VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmTaskList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Form"
   ClientHeight    =   4740
   ClientLeft      =   1815
   ClientTop       =   825
   ClientWidth     =   5625
   Icon            =   "FrmTaskList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmTaskList.frx":0CCA
   ScaleHeight     =   4740
   ScaleWidth      =   5625
   Begin MSDataListLib.DataCombo dcRuangan 
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtOrder 
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   126025731
      UpDown          =   -1  'True
      CurrentDate     =   40331
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtMasalah 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox txtKdTask 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdSimpan 
      Appearance      =   0  'Flat
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   9360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ditujukan Ke"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label txtKdChild 
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Masalah"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmTaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mdmulai As Date
Dim mdselesai As Date
Dim subKdPemeriksa() As String
Dim subJmlTotal As Integer
Dim j As Integer
Dim kedit As String

Private Sub subDcSource()
    Call msubDcSource(dcRuangan, rs, "select KdRuangan, NamaRuangan from SIM_RuanganPelaksana")
End Sub

Private Sub cmdRiawayt1_Click()

    If txtKdTask.Text = "" Then
        Call MsgBox("Simpan JOD Terlebih Dahulu", vbOKOnly, "Validasi")
    Else
        If Trim(dcStatus.Text) = "Selesai" Then
            Call MsgBox("Tidak Bisa Menambahkan Riwayat, Karena Job Sudah Selesai", vbOKOnly, "PERINGATAN")
        Else
            fratambah.Visible = True
            dcStatus2.BoundText = "02"
            dtMulaiJob.Value = Now
            dtSelesaiJob.Value = Now
            dtMulaiJob.SetFocus
        End If
    End If

End Sub

Private Sub ChkMulai_Click()
    If ChkMulai.Value = 1 Then
        dtMulaiJob.Enabled = True
        dtMulaiJob.SetFocus
    Else
        dtMulaiJob.Enabled = False

    End If
End Sub

Private Sub ChkMulai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ChkMulai.Value = 1 Then
            dtMulaiJob.Enabled = True
            dtMulaiJob.SetFocus
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub chkPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPetugas.SetFocus
    End If
End Sub

Private Sub chkSelesai_Click()
    If chkSelesai.Value = 1 Then
        dtSelesaiJob.Enabled = True
        dtSelesaiJob.SetFocus
    Else
        dtSelesaiJob.Enabled = False

    End If
End Sub

Private Sub ChkSelesai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkSelesai.Value = 1 Then
            dtSelesaiJob.Enabled = True
            dtSelesaiJob.SetFocus
        Else
            txtPetugas.SetFocus
        End If
    End If
End Sub

Private Sub cmdBatalRiwayat_Click()
    fgPerawatPerPelayanan.Refresh
    dtMulaiJob.Value = Now
    dtSelesaiJob.Value = Now
    txtPetugas.Text = ""
    txtSolusi.Text = ""
    fgPerawatPerPelayanan.clear

End Sub

Private Sub cmdBatal_Click()
    dcRuangan.Text = ""
    txtMasalah.Text = ""
End Sub

Private Sub cmdSimpan_Click()
    If txtMasalah.Text = "" Then
        Call MsgBox("Masalah Belum di Isi", vbOKOnly, "Validasi")
    Else
        If sp_Task() = False Then Exit Sub
    End If
End Sub

Private Function sp_Task() As Boolean
    On Error GoTo errLoad
    sp_StrukTerima = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        If Len(Trim(txtKdTask.Text)) = 0 Then
            .Parameters.Append .CreateParameter("KdTask", adChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("KdTask", adChar, adParamInput, 10, txtKdTask.Text)
        End If
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdPelapor", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Masalah", adChar, adParamInput, 3000, txtMasalah.Text)
        .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtOrder.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdKategory", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("KdPrioritas", adChar, adParamInput, 2, "02")
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("KdRuanganPelaksana", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("KdTingkatKesulitan", adChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("Requestke", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglMulai", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("TglSelesai", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("OutputKdTask", adChar, adParamOutput, 10, Null)
        .ActiveConnection = dbConn
        .CommandText = "SIMAdd_Task"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_Task = False
        Else

            txtKdTask.Text = .Parameters("OutputKdTask").Value
            Call MsgBox("Data Sudah Tersimpan", vbOKOnly, "PERHATIAN")
            cmdSimpan.Enabled = False

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_Task = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Function sp_Job() As Boolean
    On Error GoTo errLoad
    sp_Job = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        If Len(Trim(TxtJobIdent.Text)) = 0 Then
            .Parameters.Append .CreateParameter("KdJob", adVarChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("KdJob", adVarChar, adParamInput, 10, TxtJobIdent.Text)
        End If
        .Parameters.Append .CreateParameter("KdTask", adVarChar, adParamInput, 10, txtKdTask.Text)
        .Parameters.Append .CreateParameter("tglMulai", adDate, adParamInput, , Format(dtMulaiJob.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("tglSelesai", adDate, adParamInput, , Format(dtSelesaiJob.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Solusi", adVarChar, adParamInput, 3000, txtSolusi.Text)
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, dcStatus2.BoundText)
        .Parameters.Append .CreateParameter("Proses", adInteger, adParamInput, , Val(txtproses1.Text))
        .Parameters.Append .CreateParameter("OutputKdJobTemp", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "SIMAU_Job"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_Job = False
        Else
            TxtJobIdent.Text = .Parameters("OutputKdJobTemp").Value
            Call MsgBox("Data Sudah Tersimpan", vbOKOnly, "PERHATIAN")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_Job = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Sub cmdTutup_Click()
    Unload Me
    frmJobList.Show
End Sub

Private Sub dcPetugas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dgKategori.SetFocus
    End If
End Sub

Private Sub dcPrioritas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSimpan.SetFocus
    End If
End Sub

Private Sub dcRuang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dcPetugas.SetFocus
    End If
End Sub

Private Sub dcStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dcPrioritas.SetFocus
    End If
End Sub

Private Sub dcStatus2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtproses1.SetFocus
    End If
End Sub

Private Sub dgKategori_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dcPrioritas.SetFocus
    End If
End Sub

Private Sub dtMulaiJob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtSelesaiJob.SetFocus
    End If
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMasalah.SetFocus
    End If
End Sub

Private Sub dtselesaijob_Change()
    If dtSelesaiJob.Value < dtMulaiJob.Value Then
        Call MsgBox("Yek Ngisi Tanggal ojo Ngawur", vbOKOnly, "Palidasi")
    End If
End Sub

Private Sub dtSelesaiJob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtPetugas.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    subDcSource
    dtOrder.Value = Now
End Sub

Private Sub txtMasalah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSimpan.SetFocus
    End If
End Sub

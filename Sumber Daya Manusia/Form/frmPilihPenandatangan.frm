VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPilihPenandatangan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ControlBox      =   0   'False
   Icon            =   "frmPilihPenandatangan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5775
   Begin VB.Frame framSubInstalasi 
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   5775
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdPilih 
         Caption         =   "&Lanjutkan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dcJabatan 
         Height          =   330
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tanda Tangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   420
         Width           =   1185
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPilihPenandatangan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3960
      Picture         =   "frmPilihPenandatangan.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPilihPenandatangan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmPilihPenandatangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    Unload Me
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub cmdPilih_Click()
    If dcJabatan.BoundText = "" Then MsgBox "Silahkan pilih", vbCritical, Konfirmasi: Exit Sub
    strTandaTangan = dcJabatan.BoundText
    '//yayang.agus 2014-08-14
    strSQL = "select NamaLengkap, NamaPangkat, NIP from V_FooterPegawai where KdJabatan='" & strTandaTangan & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        MsgBox "Tidak ada data " + dcJabatan.Text, vbExclamation, "Validasi"
        frmRiwayatPegawai.Enabled = True
        Exit Sub
    End If
    '//
    Unload Me
    strSQL = "select * from V_CetakSuratPerjalananDinasPegawai where idpegawai='" & mstrIdPegawai & "' AND NoUrut='" & strNoUrut & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        MsgBox "Tidak ada data, silahkan lengkapi kelengkapan data pegawai bersangkutan", vbExclamation, "Validasi"
        frmRiwayatPegawai.Enabled = True
        Exit Sub
    End If
        
    frmRiwayatPegawai.Enabled = True
    frmCetakSuratKeteranganPerjalananDinas.Show
End Sub

Private Sub dcJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdPilih.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo hell
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
'    strsqlx = "select Value from SettingGlobal where Prefix in ('Penandatangan1','Penandatangan2')"
'    Call msubRecFO(rsx, strsqlx)
    

    'strSQL = "SELECT KdJabatan, NamaJabatan FROM Jabatan where KdJenisJabatan ='01' and KdJabatan in ('02001','01022')"
    strSQL = "SELECT KdJabatan, NamaJabatan FROM Jabatan where KdJenisJabatan ='01'" ' and KdJabatan in ('A01','01001','A02') order by namajabatan" '//yayang.agus 2014-08-12
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcJabatan.RowSource = rs
    dcJabatan.BoundColumn = "KdJabatan"
    dcJabatan.ListField = "NamaJabatan"
    Exit Sub
hell:

End Sub

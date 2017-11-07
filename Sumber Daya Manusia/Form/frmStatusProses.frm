VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatusProses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status Proses"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmStatusProses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5235
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5055
      Begin MSComctlLib.ProgressBar pgbStatus 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Ambil PIN"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lblPersen 
         Alignment       =   1  'Right Justify
         Caption         =   "100%"
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblPIN 
         AutoSize        =   -1  'True
         Caption         =   "<pin>"
         Height          =   195
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   390
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
      Left            =   3360
      Picture         =   "frmStatusProses.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStatusProses.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmStatusProses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
    Select Case strStatusSekarang
        Case "simpan"
            resetInteger = True
            frmPINAbsensiPegawai.tmrSimpan.Enabled = False
        Case "cetak"
            resetInteger = True
            frmPINAbsensiPegawai.ListView1.ListItems.clear
            frmPINAbsensiPegawai.tmrCetak.Enabled = False
        Case "prepare"
            bolPrepareFullUpload = False
        Case "upload"
            bolFullUpload = False
    End Select
    inbuff2 = ""
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Me.lblPIN.Caption = ""
End Sub

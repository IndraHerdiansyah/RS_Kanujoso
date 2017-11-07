VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmJenisDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jenis Download"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5535
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "&Download"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5295
      Begin VB.TextBox txtPinDownload 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   3375
      End
      Begin VB.OptionButton optFull 
         Caption         =   "&Full Download"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optParsial 
         Caption         =   "&Parsial Download"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
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
      Left            =   3720
      Picture         =   "frmJenisDownload.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmJenisDownload.frx":0D88
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmJenisDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
End Sub

Private Sub optFull_Click()
    If Me.optFull.Value = True Then
        Me.txtPinDownload.Enabled = False
    End If
End Sub

Private Sub optParsial_Click()
    If Me.optParsial.Value = True Then
        Me.txtPinDownload.Enabled = True
        Me.txtPinDownload.SetFocus
    End If
End Sub

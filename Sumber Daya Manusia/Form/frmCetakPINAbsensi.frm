VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakPINAbsensi 
   Caption         =   "frmCetakDaftarNoAbsensi"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   Icon            =   "frmCetakPINAbsensi.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   6315
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakPINAbsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New CrPINAbsensi

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command

    Call openConnection
    Set frmPINAbsensiPegawai = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "SELECT     ID, Nama, Jabatan, Ruangan, PIN, [Tgl. Daftar], [Alamat FRS] " & _
    " From v_PIN "

    adocomd.CommandType = adCmdUnknown
    Report.Database.AddADOCommand dbConn, adocomd
    With Report
        .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
        .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
        .txtAlamatRS.SetText strNAlamatRS
        .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos : " & " " & strNKodepos & " "

        .usIdPegawai.SetUnboundFieldSource ("{Ado.ID}")
        .usPIN.SetUnboundFieldSource ("{Ado.PIN}")
        .usAlamatFRS.SetUnboundFieldSource ("{Ado.Alamat FRS}")
        .usNamaPegawai.SetUnboundFieldSource ("{Ado.Nama}")
        .usJabatan.SetUnboundFieldSource ("{Ado.Jabatan}")
        .usRuangan.SetUnboundFieldSource ("{Ado.Ruangan}")

    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .EnableGroupTree = False
        .Zoom 1

    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPINAbsensiPegawai = Nothing
End Sub


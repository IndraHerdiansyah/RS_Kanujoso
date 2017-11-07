VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarPegawai 
   Caption         =   "CETAK"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakDaftarPegawai.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
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
Attribute VB_Name = "frmCetakDaftarPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crDaftarPegawai

Private Sub Command1_Click()
    On Error Resume Next
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd
    With Report
        .txtNamaRS.SetText strNNamaRS '& " " & strkelasRS & " " & strketkelasRS
        .txtKabupaten.SetText strNKotaRS
        .txtAlamatRS.SetText strNAlamatRS
        .txtInstalasi.SetText "" '"HUMAN RESOURCE DEPARTMENT"
        .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos : " & " " & strNKodepos & " "
        
        .usJenisPegawai.SetUnboundFieldSource ("{ado.JenisPegawai}")
        .usIdPegawai.SetUnboundFieldSource ("{ado.IdPegawai}")
        .usNamaLengkap.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .usSex.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .usTempatLahir.SetUnboundFieldSource ("{ado.TempatLahir}")
        .udTglLahir.SetUnboundFieldSource ("{ado.TglLahir}")
        .usPangkat.SetUnboundFieldSource ("{ado.NamaPangkat}")
        .usGolongan.SetUnboundFieldSource ("{ado.NamaGolongan}")
        .usJabatan.SetUnboundFieldSource ("{ado.NamaJabatan}")
        .usPendidikan.SetUnboundFieldSource ("{ado.Pendidikan}")
        .usNIP.SetUnboundFieldSource ("{ado.NIP}")
        .usStatus.SetUnboundFieldSource ("{ado.StatusAktif}")
    End With
    
    Screen.MousePointer = vbHourglass
    If vLaporan = "Print" Then
        Report.PrintOut False
        Unload Me
    Else
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 1
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakDaftarPegawai = Nothing
End Sub

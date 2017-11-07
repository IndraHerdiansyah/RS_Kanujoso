VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakRekapKinerja 
   Caption         =   "frmCetakAbsensiPegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakRekapKinerja.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
Attribute VB_Name = "frmCetakRekapKinerja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crRekapKinerja
Private Sub Command1_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Dim adocomd As New ADODB.Command

    Set frmCetakRekapKinerja = Nothing
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    With Report

        .txtNamaRS.SetText strNNamaRS
        .txtAlamatRS.SetText strNAlamatRS
        .txtTelpRS.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS

        .txtPeriode.SetText "Periode : " & Format(mdTglAwal, "MMMM") & " s/d " & Format(mdTglAkhir, "MMMM") & ""
       
'        .txtNamaRuangan.SetText mstrNamaRuangan
    End With

    Set dbcmd = New ADODB.Command
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, dbcmd
    With Report
       .usNama.SetUnboundFieldSource ("{ado.NamaLengkap}")
       .usNIlai.SetUnboundFieldSource ("{ado.NilaiTotal}")
       .usPrestasi.SetUnboundFieldSource ("{ado.NilaiPrestasi}")

        
        
    End With
    If vLaporan = "view" Then
    With CRViewer1
        .ReportSource = Report
        .EnableGroupTree = False
        .ViewReport
        .Zoom 1
    End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakRekapKinerja = Nothing
End Sub


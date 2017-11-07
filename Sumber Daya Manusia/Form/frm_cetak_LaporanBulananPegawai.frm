VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_LaporanBulananPegawai 
   Caption         =   "Cetak Laporan Bulanan Jumlah Pegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm_cetak_LaporanBulananPegawai.frx":0000
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
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frm_cetak_LaporanBulananPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crLaporanBulananJumlahPegawai

Private Sub Form_Load()
    On Error GoTo hell
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Set frm_cetak_LaporanBulananPegawai = Nothing
    Dim adocomd As New ADODB.Command

    adocomd.ActiveConnection = dbConn

    adocomd.CommandText = strSQL

    adocomd.CommandType = adCmdText

    With Report
        .Database.AddADOCommand dbConn, adocomd

        .usNamaRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .unSD.SetUnboundFieldSource ("{ado.SD}")
        .unSMP.SetUnboundFieldSource ("{ado.SMP}")
        .unSMA.SetUnboundFieldSource ("{ado.SMA}")
        .unD3.SetUnboundFieldSource ("{ado.D3}")
        .unS1.SetUnboundFieldSource ("{ado.S1}")
        .unS2.SetUnboundFieldSource ("{ado.S2}")
        .unLain2.SetUnboundFieldSource ("{ado.Lain2}")
        .unJmlTPHL.SetUnboundFieldSource ("{ado.SD} + {ado.SMP} + {ado.SMA} + {ado.D3} + {ado.S1} + {ado.S2} + {ado.Lain2}")

        .unC1a.SetUnboundFieldSource ("{ado.CPNS1a}")
        .unC1c.SetUnboundFieldSource ("{ado.CPNS1c}")
        .unC2a.SetUnboundFieldSource ("{ado.CPNS2a}")
        .unC2b.SetUnboundFieldSource ("{ado.CPNS2b}")
        .unC2c.SetUnboundFieldSource ("{ado.CPNS2c}")
        .unC3a.SetUnboundFieldSource ("{ado.CPNS3a}")
        .unC3b.SetUnboundFieldSource ("{ado.CPNS3b}")
        .unJmlCPNS.SetUnboundFieldSource ("{ado.CPNS1a} + {ado.CPNS1c} + {ado.CPNS2a} + {ado.CPNS2b} + {ado.CPNS2c} + {ado.CPNS3a} + {ado.CPNS3b}")

        .unP1a.SetUnboundFieldSource ("{ado.PNS1a}")
        .unP1b.SetUnboundFieldSource ("{ado.PNS1b}")
        .unP1c.SetUnboundFieldSource ("{ado.PNS1c}")
        .unP1d.SetUnboundFieldSource ("{ado.PNS1d}")
        .unP2a.SetUnboundFieldSource ("{ado.PNS2a}")
        .unP2b.SetUnboundFieldSource ("{ado.PNS2b}")
        .unP2c.SetUnboundFieldSource ("{ado.PNS2c}")
        .unP2d.SetUnboundFieldSource ("{ado.PNS2d}")
        .unP3a.SetUnboundFieldSource ("{ado.PNS3a}")
        .unP3b.SetUnboundFieldSource ("{ado.PNS3b}")
        .unP3c.SetUnboundFieldSource ("{ado.PNS3c}")
        .unP3d.SetUnboundFieldSource ("{ado.PNS3d}")
        .unP4a.SetUnboundFieldSource ("{ado.PNS4a}")
        .unP4b.SetUnboundFieldSource ("{ado.PNS4b}")
        .unP4c.SetUnboundFieldSource ("{ado.PNS4c}")
        .unP4d.SetUnboundFieldSource ("{ado.PNS4d}")
        .unP4e.SetUnboundFieldSource ("{ado.PNS4e}")
        .unJmlPNS.SetUnboundFieldSource ("{ado.PNS1a} + {ado.PNS1b} + {ado.PNS1c} + {ado.PNS1d} + {ado.PNS2a} + {ado.PNS2b} + {ado.PNS2c} + {ado.PNS2d} + {ado.PNS3a} + {ado.PNS3b} + {ado.PNS3c} + {ado.PNS3d} + {ado.PNS4a} + {ado.PNS4b} + {ado.PNS4c} + {ado.PNS4d} + {ado.PNS4e}")
        .txtBulan.SetText Format(Now, "MMMM yyyy")
    End With
    strSQLsplakuk = "select NamaLengkap, NamaPangkat, NIP from V_FooterPegawai where KdJabatan='A01'"
    Call msubRecFO(rsSplakuk, strSQLsplakuk)

    Report.txtDirektur.SetText IIf(IsNull(rsSplakuk("NamaLengkap").Value), "", rsSplakuk("NamaLengkap").Value)
    Report.txtNIPF.SetText IIf(IsNull(rsSplakuk("NIP").Value), "", rsSplakuk("NIP").Value)
    Screen.MousePointer = vbHourglass

    If vLaporan = "Print" Then
        Report.PrintOut False
        Unload Me
    Else

        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 100
        End With
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_cetak_LaporanBulananPegawai = Nothing
End Sub

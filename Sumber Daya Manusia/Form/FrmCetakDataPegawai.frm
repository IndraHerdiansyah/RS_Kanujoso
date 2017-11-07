VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakDataPegawai 
   Caption         =   "Medifirst2000 - Data Pegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmCetakDataPegawai.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
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
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmCetakDataPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New CrPegawai

Private Sub Form_Load()
    On Error GoTo hell
    Dim adocomd As New ADODB.Command

    Call openConnection
    Set FrmCetakDataPegawai = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn

    adocomd.CommandText = "SELECT     IdPegawai, JenisPegawai, NamaLengkap, JenisKelamin, TempatLahir, TglLahir, NamaPangkat, NamaGolongan, NamaJabatan, Pendidikan, NIP, StatusAktif " & _
    " From V_Data_Pegawai1 where kdstatus= '" & frmDataPegawaiNew.dcParamStatus.BoundText & "' and NamaLengkap LIKE '%" & frmDataPegawaiNew.txtParameter.Text & "%'"

    adocomd.CommandType = adCmdUnknown
    Report.Database.AddADOCommand dbConn, adocomd
    With Report
        .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
        .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
        .txtAlamatRS.SetText strNAlamatRS
        .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos : " & " " & strNKodepos & " "

        .usIdPegawai.SetUnboundFieldSource ("{Ado.IdPegawai}")
        .usJenisPegawai.SetUnboundFieldSource ("{Ado.JenisPegawai}")
        .UsNama.SetUnboundFieldSource ("{Ado.NamaLengkap}")
        .UsJK.SetUnboundFieldSource ("{Ado.JenisKelamin}")
        .usTempatLahir.SetUnboundFieldSource ("{Ado.TempatLahir}")
        .UsTanggallahir.SetUnboundFieldSource ("{Ado.Tgllahir}")
        .usPangkat.SetUnboundFieldSource ("{Ado.namaPangkat}")
        .usGolongan.SetUnboundFieldSource ("{Ado.namaGolongan}")
        .usJabatan.SetUnboundFieldSource ("{Ado.NamaJabatan}")
        .usPendidikan.SetUnboundFieldSource ("{Ado.Pendidikan}")
        .usNIP.SetUnboundFieldSource ("{Ado.NIP}")
        .usStatus.SetUnboundFieldSource ("{Ado.StatusAktif}")

    End With
    Screen.MousePointer = vbHourglass

'    With CRViewer1
'        .ReportSource = Report
'        .ViewReport
'        .EnableGroupTree = True
'        .Zoom 1
'    End With
   If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 1
            .DisplayGroupTree = False
        End With
    Else
        Dim tempPrint1 As String
        Dim strDeviceName As String
        Dim strDriverName As String
        Dim strPort As String
        Dim p As Printer
        Dim Posisi, z, Urutan As Integer
        
        Dim sPrinter1 As String
            
            sPrinter1 = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "PrinterLegal")
            
                Urutan = 0
                For z = 1 To Len(sPrinter1)
                    If Mid(sPrinter1, z, 1) = ";" Then
                        Urutan = Urutan + 1
                        Posisi = z
                        ReDim Preserve arrPrinter(Urutan)
                        arrPrinter(Urutan).intUrutan = Urutan
                        arrPrinter(Urutan).intPosisi = Posisi
                        If Urutan = 1 Then
                            arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter1, 1, z - 1)
                        Else
                            arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter1, arrPrinter(Urutan - 1).intPosisi + 1, z - arrPrinter(Urutan - 1).intPosisi - 1)
                        End If
                     
                     
                    For Each p In Printers
                            strDeviceName = arrPrinter(Urutan).strNamaPrinter
                            strDriverName = p.DriverName
                            strPort = p.Port
                
                            Report.SelectPrinter strDriverName, strDeviceName, strPort
                            Report.PrintOut False
                            Screen.MousePointer = vbDefault
        
                    Exit For
                    
                    Next
                End If
            Next z
              Unload Me
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
    Set FrmCetakDataPegawai = Nothing
End Sub

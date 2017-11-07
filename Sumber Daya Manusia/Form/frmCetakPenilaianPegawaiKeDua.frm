VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakPenilaianPegawaiKeDua 
   Caption         =   "Form Cetak Penilaian Pegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakPenilaianPegawaiKeDua.frx":0000
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
Attribute VB_Name = "frmCetakPenilaianPegawaiKeDua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crPenilaianPegawaiKeDua

Private Sub Form_Load()
    On Error GoTo hell
    Me.WindowState = 2
    With Report
        .txtKesetiaan.SetText IIf(IsNull(rs.Fields("NilaiKesetiaan").Value), "", rs.Fields("NilaiKesetiaan").Value)
        strsqlx = "Select Keterangan from DaftarNilai where Nilai='" & rs.Fields("NilaiKesetiaan").Value & "'"
        Call msubRecFO(rsx, strsqlx)
        
        
        .txtPrestasi.SetText IIf(IsNull(rs.Fields("NilaiPrestasi").Value), "", rs.Fields("NilaiPrestasi").Value)
        .txtTanggung.SetText IIf(IsNull(rs.Fields("NilaiTanggungJawab").Value), "", rs.Fields("NilaiTanggungJawab").Value)
        .txtKetaatan.SetText IIf(IsNull(rs.Fields("NilaiKetaatan").Value), "", rs.Fields("NilaiKetaatan").Value)
        .txtKejujuran.SetText IIf(IsNull(rs.Fields("NilaiKejujuran").Value), "", rs.Fields("NilaiKejujuran").Value)
        .txtKerjasama.SetText IIf(IsNull(rs.Fields("NilaiKerjasama").Value), "", rs.Fields("NilaiKerjasama").Value)
        .txtPrakarsa.SetText IIf(IsNull(rs.Fields("NilaiPrakarsa").Value), "", rs.Fields("NilaiPrakarsa").Value)
        .txtKepemimpinan.SetText IIf(IsNull(rs.Fields("NilaiKepemimpinan").Value), "", rs.Fields("NilaiKepemimpinan").Value)
        .txtJumlah.SetText (Val(rs.Fields("NilaiKesetiaan")) + Val(rs.Fields("NilaiPrestasi")) + Val(rs.Fields("NilaiTanggungJawab")) + Val(rs.Fields("NilaiKetaatan")) + Val(rs.Fields("NilaiKejujuran")) + Val(rs.Fields("NilaiKerjasama")) + Val(rs.Fields("NilaiPrakarsa")) + Val(rs.Fields("NilaiKepemimpinan")))
        .txtSebutanJml.SetText ""
        .txtRata.SetText (Val(rs.Fields("NilaiKesetiaan")) + Val(rs.Fields("NilaiPrestasi")) + Val(rs.Fields("NilaiTanggungJawab")) + Val(rs.Fields("NilaiKetaatan")) + Val(rs.Fields("NilaiKejujuran")) + Val(rs.Fields("NilaiKerjasama")) + Val(rs.Fields("NilaiPrakarsa")) + Val(rs.Fields("NilaiKepemimpinan"))) / 8
        .txtSebutanRata.SetText ""
        
        
        
'        If rs.Fields("NilaiKesetiaan").Value >= "91" Then
'            .txtSebutan.SetText "Amat Baik"
'        ElseIf rs.Fields("NilaiKesetiaan").Value < "91" Then
'            .txtSebutan.SetText "Baik"
'        Else
'            .txtSebutan.SetText "-"
'        End If
        If rs.Fields("NilaiKesetiaan").Value = "100" Then
            .txtSebutan.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiKesetiaan").Value >= "91" And rs.Fields("NilaiKesetiaan").Value <= 99 Then
            .txtSebutan.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiKesetiaan").Value >= "80" And rs.Fields("NilaiKesetiaan").Value <= "90" Then
            .txtSebutan.SetText "Baik"
        ElseIf rs.Fields("NilaiKesetiaan").Value >= "60" And rs.Fields("NilaiKesetiaan").Value <= "79" Then
            .txtSebutan.SetText "Cukup"
        ElseIf rs.Fields("NilaiKesetiaan").Value >= "40" And rs.Fields("NilaiKesetiaan").Value <= "59" Then
            .txtSebutan.SetText "Kurang"
        ElseIf rs.Fields("NilaiKesetiaan").Value >= "0" And rs.Fields("NilaiKesetiaan").Value <= "39" Then
            .txtSebutan.SetText "Sangat Kurang"
        Else
            .txtSebutan.SetText "-"
        End If
        
'        If rs.Fields("NilaiPrestasi").Value >= "91" Then
'            .txtSebutanPrestasi.SetText "Amat Baik"
'        ElseIf rs.Fields("NilaiPrestasi").Value < "91" Then
'            .txtSebutanPrestasi.SetText "Baik"
'        Else
'            .txtSebutanPrestasi.SetText "-"
'        End If
        If rs.Fields("NilaiPrestasi").Value = "100" Then
            .txtSebutan.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiPrestasi").Value >= "91" And rs.Fields("NilaiPrestasi").Value >= "99" Then
            .txtSebutanPrestasi.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiPrestasi").Value >= "80" And rs.Fields("NilaiPrestasi").Value <= "90" Then
            .txtSebutanPrestasi.SetText "Baik"
        ElseIf rs.Fields("NilaiPrestasi").Value >= "60" And rs.Fields("NilaiPrestasi").Value <= "79" Then
            .txtSebutanPrestasi.SetText "Cukup"
        ElseIf rs.Fields("NilaiPrestasi").Value >= "40" And rs.Fields("NilaiPrestasi").Value <= "59" Then
            .txtSebutanPrestasi.SetText "Kurang"
        ElseIf rs.Fields("NilaiPrestasi").Value >= "0" And rs.Fields("NilaiPrestasi").Value <= "39" Then
            .txtSebutanPrestasi.SetText "Sangat Kurang"
        Else
            .txtSebutanPrestasi.SetText "-"
        End If

'        If rs.Fields("NilaiTanggungJawab").Value >= "91" Then
'            .txtSebutantanggungJawab.SetText "Amat Baik"
'        ElseIf rs.Fields("NilaiTanggungJawab").Value < "91" Then
'            .txtSebutantanggungJawab.SetText "Baik"
'        Else
'            .txtSebutantanggungJawab.SetText "-"
'        End If
        If rs.Fields("NilaiTanggungJawab").Value = "100" Then
            .txtSebutantanggungJawab.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiTanggungJawab").Value = "100" Then
            .txtSebutantanggungJawab.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiTanggungJawab").Value >= "91" And rs.Fields("NilaiTanggungJawab").Value >= "99" Then
            .txtSebutantanggungJawab.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiTanggungJawab").Value >= "80" And rs.Fields("NilaiTanggungJawab").Value <= "90" Then
            .txtSebutantanggungJawab.SetText "Baik"
        ElseIf rs.Fields("NilaiTanggungJawab").Value >= "60" And rs.Fields("NilaiTanggungJawab").Value <= "79" Then
            .txtSebutantanggungJawab.SetText "Cukup"
        ElseIf rs.Fields("NilaiTanggungJawab").Value >= "40" And rs.Fields("NilaiTanggungJawab").Value <= "59" Then
            .txtSebutantanggungJawab.SetText "Kurang"
        ElseIf rs.Fields("NilaiTanggungJawab").Value >= "0" And rs.Fields("NilaiTanggungJawab").Value <= "39" Then
            .txtSebutantanggungJawab.SetText "Sangat Kurang"
        Else
            .txtSebutantanggungJawab.SetText "-"
        End If

'        If rs.Fields("NilaiKetaatan").Value >= "91" Then
'            .txtSebutanKetaatan.SetText "Amat Baik"
'        ElseIf rs.Fields("NilaiKetaatan").Value < "91" Then
'            .txtSebutanKetaatan.SetText "Baik"
'        Else
'            .txtSebutanKetaatan.SetText "-"
'        End If
        If rs.Fields("NilaiKetaatan").Value = "100" Then
            .txtSebutanKetaatan.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiKetaatan").Value >= "91" And rs.Fields("NilaiKetaatan").Value >= "99" Then
            .txtSebutanKetaatan.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiKetaatan").Value >= "80" And rs.Fields("NilaiKetaatan").Value <= "90" Then
            .txtSebutanKetaatan.SetText "Baik"
        ElseIf rs.Fields("NilaiKetaatan").Value >= "60" And rs.Fields("NilaiKetaatan").Value <= "79" Then
            .txtSebutanKetaatan.SetText "Cukup"
        ElseIf rs.Fields("NilaiKetaatan").Value >= "40" And rs.Fields("NilaiKetaatan").Value <= "59" Then
            .txtSebutanKetaatan.SetText "Kurang"
        ElseIf rs.Fields("NilaiKetaatan").Value >= "1" And rs.Fields("NilaiKetaatan").Value <= "39" Then
            .txtSebutanKetaatan.SetText "Sangat Kurang"
        Else
            .txtSebutanKetaatan.SetText "-"
        End If

'        If rs.Fields("NilaiKejujuran").Value >= "91" Then
'            .txtSebutanKejujuran.SetText "Amat Baik"
'        ElseIf rs.Fields("NilaiKejujuran").Value < "91" Then
'            .txtSebutanKejujuran.SetText "Baik"
'        Else
'            .txtSebutanKejujuran.SetText "-"
'        End If
        If rs.Fields("NilaiKejujuran").Value = "100" Then
            .txtSebutanKejujuran.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiKejujuran").Value >= "91" And rs.Fields("NilaiKejujuran").Value >= "99" Then
            .txtSebutanKejujuran.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiKejujuran").Value >= "80" And rs.Fields("NilaiKejujuran").Value <= "90" Then
            .txtSebutanKejujuran.SetText "Baik"
        ElseIf rs.Fields("NilaiKejujuran").Value >= "60" And rs.Fields("NilaiKejujuran").Value <= "79" Then
            .txtSebutanKejujuran.SetText "Cukup"
        ElseIf rs.Fields("NilaiKejujuran").Value >= "40" And rs.Fields("NilaiKejujuran").Value <= "59" Then
            .txtSebutanKejujuran.SetText "Kurang"
        ElseIf rs.Fields("NilaiKejujuran").Value >= "0" And rs.Fields("NilaiKejujuran").Value <= "39" Then
            .txtSebutanKejujuran.SetText "Sangat Kurang"
        Else
            .txtSebutanKejujuran.SetText "-"
        End If

'        If rs.Fields("NilaiKerjasama").Value >= "91" Then
'            .txtSebutanKerjasama.SetText "Amat Baik"
'        ElseIf rs.Fields("NilaiKerjasama").Value < "91" Then
'            .txtSebutanKerjasama.SetText "Baik"
'        Else
'            .txtSebutanKerjasama.SetText "-"
'        End If
        If rs.Fields("NilaiKerjasama").Value = "100" Then
            .txtSebutanKerjasama.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiKerjasama").Value >= "91" And rs.Fields("NilaiKerjasama").Value >= "99" Then
            .txtSebutanKerjasama.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiKerjasama").Value >= "80" And rs.Fields("NilaiKerjasama").Value <= "90" Then
            .txtSebutanKerjasama.SetText "Baik"
        ElseIf rs.Fields("NilaiKerjasama").Value >= "60" And rs.Fields("NilaiKerjasama").Value <= "79" Then
            .txtSebutanKerjasama.SetText "Cukup"
        ElseIf rs.Fields("NilaiKerjasama").Value >= "40" And rs.Fields("NilaiKerjasama").Value <= "59" Then
            .txtSebutanKerjasama.SetText "Kurang"
        ElseIf rs.Fields("NilaiKerjasama").Value >= "0" And rs.Fields("NilaiKerjasama").Value <= "39" Then
            .txtSebutanKerjasama.SetText "Sangat Kurang"
        Else
            .txtSebutanKerjasama.SetText "-"
        End If

'        If rs.Fields("NilaiPrakarsa").Value >= "91" Then
'            .txtSebutanPrakarsa.SetText "Amat Baik"
'        ElseIf rs.Fields("NilaiPrakarsa").Value < "91" Then
'            .txtSebutanPrakarsa.SetText "Baik"
'        Else
'            .txtSebutanPrakarsa.SetText "-"
'        End If
        If rs.Fields("NilaiPrakarsa").Value = "100" Then
            .txtSebutanPrakarsa.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiPrakarsa").Value >= "91" And rs.Fields("NilaiPrakarsa").Value >= "99" Then
            .txtSebutanPrakarsa.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiPrakarsa").Value >= "80" And rs.Fields("NilaiPrakarsa").Value <= "90" Then
            .txtSebutanPrakarsa.SetText "Baik"
        ElseIf rs.Fields("NilaiPrakarsa").Value >= "60" And rs.Fields("NilaiPrakarsa").Value <= "79" Then
            .txtSebutanPrakarsa.SetText "Cukup"
        ElseIf rs.Fields("NilaiPrakarsa").Value >= "40" And rs.Fields("NilaiPrakarsa").Value <= "59" Then
            .txtSebutanPrakarsa.SetText "Kurang"
        ElseIf rs.Fields("NilaiPrakarsa").Value >= "0" And rs.Fields("NilaiPrakarsa").Value <= "39" Then
            .txtSebutanPrakarsa.SetText "Sangat Kurang"
        Else
            .txtSebutanPrakarsa.SetText "-"
        End If

'        If rs.Fields("NilaiKepemimpinan").Value >= "91" Then
'            .txtSebutanKep.SetText "Amat Baik"
'        ElseIf rs.Fields("NilaiKepemimpinan").Value < "91" Then
'            .txtSebutanKep.SetText "Baik"
'        Else
'            .txtSebutanKep.SetText "-"
'        End If
'        .txtSebutanRata.SetText "Baik"
        If rs.Fields("NilaiKepemimpinan").Value = "100" Then
            .txtSebutanKep.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiKepemimpinan").Value >= "91" And rs.Fields("NilaiKepemimpinan").Value >= "99" Then
            .txtSebutanKep.SetText "Amat Baik"
        ElseIf rs.Fields("NilaiKepemimpinan").Value >= "80" And rs.Fields("NilaiKepemimpinan").Value <= "90" Then
            .txtSebutanKep.SetText "Baik"
        ElseIf rs.Fields("NilaiKepemimpinan").Value >= "60" And rs.Fields("NilaiKepemimpinan").Value <= "79" Then
            .txtSebutanKep.SetText "Cukup"
        ElseIf rs.Fields("NilaiKepemimpinan").Value >= "40" And rs.Fields("NilaiKepemimpinan").Value <= "59" Then
            .txtSebutanKep.SetText "Kurang"
        ElseIf rs.Fields("NilaiKepemimpinan").Value >= "0" And rs.Fields("NilaiKepemimpinan").Value <= "39" Then
            .txtSebutanKep.SetText "Sangat Kurang"
        Else
            .txtSebutanKep.SetText "-"
        End If
        '.txtSebutanRata.SetText "Baik"
        
        If .txtRata.Text = "100" Then
            .txtSebutanRata.SetText "Amat Baik"
        ElseIf .txtRata.Text >= 91 And .txtRata.Text >= 99 Then
            .txtSebutanRata.SetText "Amat Baik"
        ElseIf .txtRata.Text >= 80 And .txtRata.Text <= 90 Then
            .txtSebutanRata.SetText "Baik"
        ElseIf .txtRata.Text >= 60 And .txtRata.Text <= 79 Then
            .txtSebutanRata.SetText "Cukup"
        ElseIf .txtRata.Text >= 40 And .txtRata.Text <= 59 Then
            .txtSebutanRata.SetText "Kurang"
         ElseIf .txtRata.Text >= 0 And .txtRata.Text <= 39 Then
            .txtSebutanRata.SetText "Sangat Kurang"
        End If

       
    End With
    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 1
            .DisplayGroupTree = False
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
hell:
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakPenilaianPegawaiKeDua = Nothing
End Sub

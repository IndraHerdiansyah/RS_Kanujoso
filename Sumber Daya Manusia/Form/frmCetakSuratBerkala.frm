VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSuratBerkala 
   Caption         =   "Cetak Surat Berkala"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "frmCetakSuratBerkala.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6630
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
Attribute VB_Name = "frmCetakSuratBerkala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crSuratKeteranganBerkala
Dim GajiLama As Double
Dim Komponen As String
Public KdKomponenGaji As String

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    On Error GoTo Errload

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    
    Report.txtNamaKota.SetText strNKotaRS
    
    strNoUrut = strNoUrut ' - 1

    If Len(strNoUrut) = 1 Then
        strNoUrut = "0" + strNoUrut
    End If
    
    strSQLsplakuk = "select * from V_CetakGajiBerkala where idpegawai='" & strIDPegawai & "' and KdKomponenGaji='" & KdKomponenGaji & "'"
    Call msubRecFO(rsSplakuk, strSQLsplakuk)
    
    Dim i As Integer
    Dim noUrut As String
    For i = 0 To rsSplakuk.RecordCount - 1
        noUrut = rsSplakuk("NoUrut").Value
        rsSplakuk.MoveNext
        If (rsSplakuk.EOF) Then
            noUrut = "00000"
            Exit For
        End If
        If (rsSplakuk("NoUrut").Value = strNoUrut) Then
          Exit For
        End If
    Next i
    
    strSQLsplakuk = "select * from V_CetakGajiBerkala where idpegawai='" & strIDPegawai & "' AND NoUrut='" & noUrut & "'"
    Call msubRecFO(rsSplakuk, strSQLsplakuk)
    
        If rsSplakuk.EOF = False Then
            GajiLama = FormatCurrency(rsSplakuk.Fields("Jumlah").Value, 2)
            Komponen = rsSplakuk.Fields("KomponenGaji").Value
        Else
            GajiLama = "0"
            Komponen = rs.Fields("KomponenGaji").Value
        End If
        
        If rsSplakuk.EOF = False Then
            If rs.Fields("KomponenGaji").Value = rsSplakuk.Fields("KomponenGaji").Value Then
                GajiLama = FormatCurrency(rsSplakuk.Fields("Jumlah").Value, 2)
                Komponen = rsSplakuk.Fields("KomponenGaji").Value
            Else
                '------------------------------------------
                strSQLsplakuk = "select * from V_CetakGajiBerkala where idpegawai='" & strIDPegawai & "' AND KomponenGaji='" & rs.Fields("KomponenGaji").Value & "'"
                Call msubRecFO(rsSplakuk, strSQLsplakuk)
                 'Dim noUrut As String
'                 Dim a As Integer
'                 For a = 1 To rsSplakuk.RecordCount
'                    If a = 1 Then
''                        rsSplakuk.MoveFirst
'                        nourut = rsSplakuk.Fields("Nourut").Value
'                    End If
''                    rsSplakuk.MoveNext
'                 Next a
                 
                 If noUrut = strNoUrut Then
                    GajiLama = "0"
                    Komponen = rsSplakuk.Fields("KomponenGaji").Value
                 Else
                   
                    strNoUrut = strNoUrut - 1
                   
                    If Len(strNoUrut) = 1 Then
                        strNoUrut = "0" + strNoUrut
                    End If
                '-------------------------------------------
                    If strNoUrut = "00" Then
                         GajiLama = "0"
                         Komponen = rsSplakuk.Fields("KomponenGaji").Value
                    Else
                         GajiLama = FormatCurrency(rsSplakuk.Fields("Jumlah").Value, 2)
                         Komponen = rsSplakuk.Fields("KomponenGaji").Value
                    End If
                 End If
            End If
        End If
        
    With Report
        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        .Database.AddADOCommand dbConn, adocomd

        .txtNama.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
        .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
        .txtJabatan.SetText IIf(IsNull(rs.Fields("NamaJabatan").Value), "", rs.Fields("NamaJabatan").Value)
        .txtPangkat.SetText IIf(IsNull(rs.Fields("Pangkat Golongan").Value), "", rs.Fields("Pangkat Golongan").Value)
        If (IsNull(rs.Fields("TempatLahir").Value)) Then
            .txtLahir.SetText ""
        Else
            .txtLahir.SetText rs.Fields("TempatLahir").Value + "," + " " + Format(rs.Fields("TglLahir").Value, "dd MMMM yyyy")
        End If
        
'        .txtLama.SetText rs.Fields("KomponenGaji").Value + " " + "Lama"
'
'        If rs.RecordCount <= 1 Then
'            .txtGajiLama.SetText 0
'        Else
''            .txtGajiLama.SetText FormatCurrency(rs.Fields("Jumlah").Value, 2)
'            .txtGajiLama.SetText FormatCurrency(frmRiwayatPegawai.dgRiwayatGaji.Columns("Jumlah"), 2)
'            .txtTerbilangGajiLama.SetText NumToTextRupiah(frmRiwayatPegawai.dgRiwayatGaji.Columns("Jumlah"))
'        End If
        
        '.UsKomponenGaji.SetUnboundFieldSource ("{Ado.KomponenGaji}")
        
        If Len(strNoUrut) = strNoUrut Then
            .txtGajiLama.SetText FormatCurrency(GajiLama, 2)
            .txtLama.SetText Komponen + " " + "(Lama)"
        Else
            If Len(strNoUrut) = "00" Then
                .txtGajiLama.SetText "0"
                .txtGajiLama.SetText FormatCurrency(rs.Fields("Jumlah").Value, 2)
             Else
                .txtGajiLama.SetText FormatCurrency(GajiLama, 2)
                .txtLama.SetText Komponen + " " + "(Lama)"
             End If
        End If
    
        .txtTerbilangGajiLama.SetText "(" + NumToTextRupiah(GajiLama) + ")"
        
        .txtNamaPejabat.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .txtTglKeputusan.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", Format(rs.Fields("TglSK").Value, "dd MMMM yyyy"))
        .txtNoKeputusan.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)

        .txtTglBerlakuGaji.SetText IIf(IsNull(rs.Fields("TglBerlakuSK").Value), "", Format(rs.Fields("TglBerlakuSK").Value, "dd MMMM yyyy"))
        .txtMasaKerja.SetText IIf(IsNull(rs.Fields("MasaKerja").Value), "", rs.Fields("MasaKerja").Value)

'        .txtGajiBaru.SetText FormatCurrency(rs.Fields("Jumlah").Value, 2)
'        .txtBaru.SetText rs.Fields("KomponenGaji").Value + " " + "Baru"
'        .txtTerbilangGajiBaru.SetText "(" + NumToTextRupiah(rs.Fields("Jumlah").Value) + ")"
        
        '.UsKomponenGaji2.SetUnboundFieldSource ("{Ado.KomponenGaji}")
        
        If strNoUrut = strNoUrut Then
            .txtGajiBaru.SetText FormatCurrency(rs.Fields("Jumlah").Value, 2)
            .txtBaru.SetText IIf(IsNull(rs.Fields("KomponenGaji").Value + " " + "(Baru)"), "", rs.Fields("KomponenGaji").Value + " " + "(Baru)")
        Else
            If Len(strNoUrut) = strNoUrut Then
                .txtGajiBaru.SetText FormatCurrency(rsSplakuk.Fields("Jumlah").Value, 2)
            Else
                .txtGajiBaru.SetText FormatCurrency(rs.Fields("Jumlah").Value, 2)
                .txtBaru.SetText IIf(IsNull(rs.Fields("KomponenGaji").Value + " " + "(Baru)"), "", rs.Fields("KomponenGaji").Value + " " + "(Baru)")
            End If
        End If
        
        .txtTerbilangGajiBaru.SetText "(" + NumToTextRupiah(rs.Fields("Jumlah").Value) + ")"
        
        .txtMasaKerjaBaru.SetText IIf(IsNull(rs.Fields("MasaKerja").Value), "", rs.Fields("MasaKerja").Value)
        .txtGolonganBaru.SetText IIf(IsNull(rs.Fields("NamaGolongan").Value), "", rs.Fields("NamaGolongan").Value)
        .txtTglBaru.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", Format(rs.Fields("TglSK").Value, "dd MMMM yyyy"))
    
        strsqlx = "select NamaLengkap, NamaPangkat, NIP from V_FooterPegawai where KdJabatan='02001'"
        Call msubRecFO(rsx, strsqlx)
        
        If rsx.BOF = True Then
            .txtPenanggungJwb.Suppress = False
            .txtPenanggungJwb.SetText "Penanggung Jawab"
            'Exit Sub
        Else
            .txtDirektur.SetText IIf(IsNull(rsx("NamaLengkap").Value), "", rsx("NamaLengkap").Value)
            .txtPangkatF.SetText IIf(IsNull(rsx("NamaPangkat").Value), "", rsx("NamaPangkat").Value)
            .txtNIPF.SetText IIf(IsNull(rsx("NIP").Value), "", rsx("NIP").Value)
        End If
        
    End With
    If vLaporan = "Print" Then
        Report.PrintOut False
        Unload Me
    Else
        CRViewer1.ReportSource = Report
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

Errload:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakSuratBerkala = Nothing
End Sub

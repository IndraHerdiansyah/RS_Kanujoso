Attribute VB_Name = "modFingerPrint"
Dim fRS As String                   'frs = variabel jumlah alamat / alamat ke
Dim code, code2 As String           'code & code2 = variabel untuk kode hex pada timer minta_absensi

Private Sub Form_Load()

    mnkoneksi.Enabled = True            'mengaktifkan tombol koneksi
    mnPKoneksi.Enabled = False          'menonaktifkan tombol putus koneksi
    add2 = 0                            'add2 = variabel untuk alamat FRS yg terkoneksi
    a = 0                               'a = variabel untuk menentukan baris msflexgrid

    'setting tampilan msflexgrid
    With MSFlexGrid1
        .ColWidth(0) = 500
        .ColWidth(1) = 1100
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 2300
        .ColWidth(5) = 2300
        .Width = 9550
        .TextMatrix(0, 0) = "FRS"
        .TextMatrix(0, 1) = "ID PEGAWAI"
        .TextMatrix(0, 2) = "MASUK / KELUAR"
        .TextMatrix(0, 3) = "STATUS"
        .TextMatrix(0, 4) = "TANGGAL"
        .TextMatrix(0, 5) = "WAKTU"
    End With

End Sub

'timer untuk mengirim kode konfirmasi FRS
Private Sub kirim_konfirmasi_Timer()

    If fRS = add Then                       'apabila add(variabel alamat FRS yg dicari) = frs(variabel jumlah alat) maka :
        kirim_konfirmasi.Enabled = False    'mematikan timer kirim_konfirmasi
        a = 0                               'variabel baris msflexgrid(a) = 0 / direset
        frmKFRS.Show                        'menampilkan frmKFRS
    Else
        fRS = (fRS + &H1)                   'apabila add <> frs maka alamat alat sekarang + 1
        MSComm1.Output = Chr$(&H2) & Chr$(&H5) & Chr$(fRS) & Chr$(&H4) & Chr$(&H3) & Chr$(fRS)  'kirim kode konfirmasi ke serial
    End If

End Sub

'timer untuk mengirim kode untuk meminta hasil absensi
Private Sub minta_absensi_Timer()

    If c = 1 Then                           'jika c (kode bantu tanda untuk meminta absen kembali) = 1

        If fRS = add2 Then                      'apabila frs = add2 maka :
            fRS = &H0                           'frs direset ke alamat 0
        End If

        code2 = code                        'isi code2 = code
        fRS = (fRS + &H1)                   'alamat alat sekarang + 1
        MSComm1.Output = Chr$(&H2) & Chr$(&H5) & Chr$(fRS) & Chr$(&H1) & Chr$(&H3) & Chr$(code) 'kirim kode meminta absen ke serial

        'pemilihan kode akhir untuk minta absen
        Select Case add2
            Case 1
                code = &H4
            Case 2
                Select Case code
                    Case &H4
                        code = &H7
                    Case &H7
                        code = &H4
                End Select
            Case 3
                Select Case code
                    Case &H4
                        code = &H7
                    Case &H7
                        code = &H6
                    Case &H6
                        code = &H4
                End Select
        End Select

    ElseIf c = 2 Then                       'jika c = 2
        MSComm1.Output = Chr$(&H6) & Chr$(&H2) & Chr$(&H5) & Chr$(fRS) & Chr$(&H1) & Chr$(&H3) & Chr$(code2)    'kirim kode meminta absen kembali ke serial

    End If

End Sub

Private Sub mnexit_Click()

    End

End Sub

'perintah untuk melakukan koneksi serial
Private Sub mnkoneksi_Click()

    If add = "" Then                        'jika variabel jumlah alat tidak ditentukan maka akan tampil pesan
        MsgBox "Setting Parameter Koneksi Kurang", vbOKOnly, "Pesan koneksi"
        Exit Sub
    Else
        mnkoneksi.Enabled = False           'jika sudah ditentukan maka tombol koneksi dinonaktifkan
        mnPKoneksi.Enabled = True           'tombol putuk koneksi diaktifkan
        mnsetting.Enabled = False           'tombol setting dinonaktifkan
        Call Get_Connect                    'memanggil sub get_connect
        fRS = &H0                           'frs direset
        code = &H4                          'code direset
        a = 0                               'a direset
        c = 1                               'c direset
        mnfile.Enabled = False              'tombol file dinonaktifkan
        kirim_konfirmasi.Enabled = True     'timer kirim_konfirmasi diaktifkan
    End If

End Sub

'perintah untuk memutus koneksi serial
Private Sub mnPKoneksi_Click()

    i = MsgBox("Putus Koneksi dengan FRS-400!", vbOKCancel, "Putus koneksi")    'menampilkan pesan konfirmasi

    If i = vbOK Then                        'jika tombol OK ditekan maka :
        minta_absensi.Enabled = False       'timer minta_absensi dimatikan
        Call Get_Disconnect                 'memanggil sub get_disconnect
        mnkoneksi.Enabled = True            'tombol koneksi diaktifkan
        mnPKoneksi.Enabled = False          'tombol putus koneksi dinonaktifkan
        mnsetting.Enabled = True            'tombol setting diaktifkan
        add2 = 0                            'add2 direset
    End If

End Sub

Public Sub mnsetting_Click()

    frmSetting.Show

End Sub

Public Sub Get_Connect()

    On Error GoTo Handle_Error              'jika error pergi ke handle_error
    If MSComm1.PortOpen = True Then         'jika port serial terbuka
        MSComm1.PortOpen = False            'tutup port serial terlebih dulu
    End If
    MSComm1.PortOpen = True                 'buka kembali port serial
    Exit Sub

    'pesan error
Handle_Error:
    MsgBox Error$, 48, "Konfirmasi Kesalahan Setting"
    Get_Disconnect

End Sub

Public Sub Get_Disconnect()

    If MSComm1.PortOpen = True Then         'jika port serial terbuka
        MSComm1.PortOpen = False             'tutup port serial
    End If

End Sub

Private Sub MSComm1_OnComm()

    Dim inbuff As String                    'inbuff = variabel untuk menyimpan input serial

    Select Case MSComm1.CommEvent

            ' Errors
        Case comEventBreak                  'A Break was received.
        Case comEventCDTO                   'CD (RLSD) Timeout.
        Case comEventCTSTO                  'CTS Timeout.
        Case comEventDSRTO                  'DSR Timeout.
        Case comEventFrame                  'Framing Error.
        Case comEventOverrun                'Data Lost.
        Case comEventRxOver                 'Receive buffer overflow.
        Case comEventRxParity               'Parity Error.
        Case comEventTxFull                 'Transmit buffer full.
        Case comEventDCB                    'Unexpected error retrieving DCB

            ' Events
        Case comEvCD                        'Change in the CD line.
        Case comEvCTS                       'Change in the CTS line.
        Case comEvDSR                       'Change in the DSR line.
        Case comEvRing                      'Change in the Ring Indicator.
        Case comEvReceive                   'Received RThreshold # of chars.

            inbuff = MSComm1.Input          'menyimpan input serial ke inbuff
            inbuff2 = inbuff2 + inbuff      'menggabungkan 2 byte pertama dengan 2 byte berikutnya di inbuff2
            Text1.Text = inbuff2            'menampilkan data di inbuff2 ke text1

            If Len(inbuff2) = 21 Then       'jika panjang data inbuff2 = 21 maka :
                Call tabel                  'memanggil sub tabel
                inbuff2 = ""                'inbuff2 direset
            ElseIf Len(inbuff2) = 30 Then   'jika panjang data inbuff2 = 30 maka :
                c = 2                       'c = 2
                Call tabel2                 'memanggil sub tabel2
                inbuff2 = ""                'inbuff2 direset
            ElseIf Len(inbuff2) = 6 Then    'jika panjang data inbuff2 = 6 maka :
                c = 1                       'c = 1
                inbuff2 = ""                'inbuff2 direset
            ElseIf Len(inbuff2) > 224 Then  'jika panjang data inbuff2 = 21 maka :
                c = 1                       'c = 1
                inbuff2 = ""                'inbuff2 direset
            End If

        Case comEvSend                      'There are SThreshold number of characters in the transmit buffer.
        Case comEvEOF                       'An EOF character was found in the

    End Select

End Sub

'perintah pada tabel hasil konfirmasi FRS
Private Sub tabel()

    add2 = add2 + 1                         'alamat FRS yg terkoneksi(add2) + 1

    With frmKFRS.MSFlexGrid1
        .Rows = .Rows + 1                   'baris msflexgrid + 1
        .TextMatrix(a, 0) = "FRS - " & Hex(Asc(Mid(inbuff2, 3, 1))) 'tulis kolom 1 dg alamat FRS
        .TextMatrix(a, 1) = Mid(inbuff2, 12, 2) + "/" + Mid(inbuff2, 10, 2) + "/" + Mid(inbuff2, 6, 4)  'tulis kolom 2 dg tanggal dr FRS
        .TextMatrix(a, 2) = Mid(inbuff2, 14, 2) + ":" + Mid(inbuff2, 16, 2) + ":" + Mid(inbuff2, 18, 2) 'tulis kolom 3 dg waktu dr FRS
        a = a + 1                           'a + 1
    End With

End Sub

'perintah pada tabel hasil absensi dari FRS
Private Sub tabel2()
    With MSFlexGrid1

        a = a + 1                           'a + 1
        .Rows = .Rows + 1                   'baris msflexgrid + 1
        .TextMatrix(a, 0) = Hex(Asc(Mid(inbuff2, 3, 1)))                'tulis kolom 1 dg alamat FRS
        .TextMatrix(a, 1) = Format(Mid(inbuff2, 21, 8), "########")     'tulis kolom 2 dg ID Pegawai

        'memeriksa kode ke 5 dari inbuff2 apakah masuk / keluar dan tulis pada kolom 3
        If Mid(inbuff2, 5, 1) = 1 Then
            .TextMatrix(a, 2) = "MASUK"
        ElseIf Mid(inbuff2, 5, 1) = 2 Then
            .TextMatrix(a, 2) = "KELUAR"
        End If

        'memeriksa kode ke 6 dari inbuff2 apakah gagal / sukses dan tulis pada kolom 4
        If Mid(inbuff2, 6, 1) = 0 Then
            .TextMatrix(a, 3) = "GAGAL"
        ElseIf Mid(inbuff2, 6, 1) = 1 Then
            .TextMatrix(a, 3) = "SUKSES"
        End If

        .TextMatrix(a, 4) = Mid(inbuff2, 13, 2) + "/" + Mid(inbuff2, 11, 2) + "/" + Mid(inbuff2, 7, 4)  'tulis kolom 5 dg tanggal dr FRS
        .TextMatrix(a, 5) = Mid(inbuff2, 15, 2) + ":" + Mid(inbuff2, 17, 2) + ":" + Mid(inbuff2, 19, 2) 'tulis kolom 6 dg waktu dr FRS

    End With

End Sub

'=============== TAMBAHAN 9 Januari 2008 ================================
Public Function cekSum(ByVal sData As String) As String
    Dim i As Integer, lData As Integer
    Dim k As String, kode As Integer, sum As Integer

    lData = Len(sData)
    For i = 1 To lData
        k = Mid(sData, i, 1)
        kode = Asc(k)
        sum = sum Xor kode
    Next

    cekSum = sum
End Function

Public Function ValidPIN(ByVal pin As String) As String
    Dim nChr As Integer
    Dim vPin As String

    If Len(pin) < 8 Then
        nChr = 8 - Len(pin)
        vPin = String(nChr, "0") & pin
    End If
    ValidPIN = vPin
End Function

Private Function HexKeDesimal(ByVal heksa As String) As Integer
    Dim i As Integer, N As Integer
    Dim h As String, d As Integer, d2 As Integer

    N = 1
    For i = 1 To 2
        h = UCase$(Mid(heksa, i, 1))

        Select Case h
            Case "A"
                d = 10
            Case "B"
                d = 11
            Case "C"
                d = 12
            Case "D"
                d = 13
            Case "E"
                d = 14
            Case "F"
                d = 15
            Case Else
                d = CInt(h)
        End Select
        d2 = d2 + ((16 ^ N) * d)
        N = N - 1
    Next
    HexKeDesimal = d2
End Function

Public Function ImageHexKeAscii(ByVal img As String) As String
    Dim i As Integer, l As Integer
    Dim des As Integer, heksa As String
    Dim imgAscii As String

    l = Len(img)
    For i = 1 To l Step 2
        heksa = Mid(img, i, 2)
        des = HexKeDesimal(heksa)
        imgAscii = imgAscii & Chr$(des)
    Next
    ImageHexKeAscii = imgAscii
End Function

'=========================================================================
'=============== TAMBAHAN 29 Januari 2008 ================================
Public Sub subCekJumlahPIN(ByVal alamatFRS As String)
    If alamatFRS = "" Then
        MsgBox "Alamat FRS Kosong!", vbCritical, "Perhatian"
        Exit Sub
    End If
    frmAbsensiPegawai.minta_absensi.Enabled = False
    frsTujuan = alamatFRS
    protokolFullUpload = Chr$(&H2) & Chr$(&H5) & Chr$(CInt(frsTujuan)) & _
    Chr$(&H13) & Chr$(&H3)
    protokolFullUpload = protokolFullUpload & Chr$(cekSum(protokolFullUpload))

    bolCekJumlahPIN = True
    frmAbsensiPegawai.MSComm1.Output = protokolFullUpload
End Sub

Public Function funcPrepareFullUploadProtocol(ByVal nData As String) As String
    Dim lData As Integer
    Dim strDum As String, validNdata As String

    lData = Len(nData)
    If lData < 4 Then
        strDum = String(4 - lData, "0")
        validNdata = strDum & nData
    End If

    protokolPrepareFullUpload = Chr$(&H2) & Chr$(&H9) & Chr$(CInt(frsTujuan)) & _
    Chr$(&H13) & validNdata & Chr$(&H3)
    protokolPrepareFullUpload = protokolPrepareFullUpload & Chr$(cekSum(protokolPrepareFullUpload))

    funcPrepareFullUploadProtocol = protokolPrepareFullUpload
End Function

'=========================================================================
'=============== TAMBAHAN 30 Januari 2008 ================================
Public Sub subAmbilPIN(ByVal dataPrepareFullUpload As String, ByVal pgb As ProgressBar, ByVal lblStat As Label, ByVal lblPIN As Label)
    Dim strData As String
    Dim keyData As Integer, idxStartPIN As Integer, panjangDataPIN As Integer
    Dim intJumlahPIN As Integer, i As Integer, j As Integer, k As Integer
    Dim pinMentah As String, pinBener As String
    Dim bilHex As String

    strData = dataPrepareFullUpload
    keyData = InStr(1, strData, Chr$(&H13), vbTextCompare)
    If keyData > 0 Then
        pgb.Max = jumlahTotalPIN
        pgb.Min = 0
        lblStat.Caption = "Ambil PIN"
        jumlahPIN = Mid(strData, keyData + 1, 4)
        intJumlahPIN = CInt(jumlahPIN)
        idxStartPIN = keyData + 5
        '        j = 0
        panjangDataPIN = Len(strData) - 2
        For i = 9 To panjangDataPIN Step 4
            '            j = j + 1
            pinBener = ""
            pinMentah = Mid(strData, i, 4)
            For k = 1 To 4
                bilHex = Hex$(Asc(Mid(pinMentah, k, 1)))
                If Len(bilHex) = 1 Then bilHex = "0" & bilHex
                pinBener = pinBener & bilHex
            Next
            idxTempDataPIN = idxTempDataPIN + 1
            tempDataPIN(idxTempDataPIN) = CInt(pinBener)
            lblPIN.Caption = CInt(pinBener)
            pgb.Value = idxTempDataPIN
        Next
    End If
End Sub

Public Sub subCetakPINKeList(ByVal lsv As ListView, ByVal pgb As ProgressBar, ByVal lblStat As Label, ByVal lblPIN As Label)
    Dim i As Integer, N As Integer, unID As Integer, j As Integer
    Dim cr As Integer, sama As Integer
    '    Dim ls As New Scripting.FileSystemObject

    '    lsv.ListItems.clear

    lblStat.Caption = "Cetak PIN"
    pgb.Max = jumlahTotalPIN
    pgb.Min = 0
    pgb.Value = 0
    For i = 1 To jumlahTotalPIN
        '        rs = Nothing
        lblPIN.Caption = tempDataPIN(i)
        strSQL = "SELECT * FROM v_PIN WHERE PIN=" & "'" & tempDataPIN(i) & "'"
        dbConn.Execute strSQL
        Call msubRecFO(rs, strSQL)
        If lsv.ListItems.Count = 0 Then
            If rs.RecordCount = 0 Then
                'If tempDataPIN(i) = "" Then GoTo lanjutkan
                '            frmPercobaan.lstPIN.AddItem tempDataPIN(i)
                lsv.ListItems.add(, , tempDataPIN(i)).SubItems(1) = frsTujuan
                unID = unID + 1
                '            ls.SubItems(0) = tempDataPIN(i)
            Else
                With lsv.ListItems.add(, , tempDataPIN(i))
                    If IsNull(rs.Fields.Item("Alamat FRS").Value) Then
                        .SubItems(1) = frsTujuan
                    Else
                        cr = InStr(1, rs.Fields.Item("Alamat FRS").Value, frsTujuan, vbTextCompare)
                        If cr = 0 Then
                            .SubItems(1) = rs.Fields.Item("Alamat FRS").Value & "," & frsTujuan
                        Else
                            .SubItems(1) = rs.Fields.Item("Alamat FRS").Value
                        End If
                    End If
                    .SubItems(2) = IIf(IsNull(rs.Fields.Item("Nama").Value), "", rs.Fields.Item("Nama").Value)
                    .SubItems(3) = IIf(IsNull(rs.Fields.Item("JK").Value), "", rs.Fields.Item("JK").Value)
                    .SubItems(4) = IIf(IsNull(rs.Fields.Item("ID").Value), "", rs.Fields.Item("ID").Value)
                    .SubItems(5) = IIf(IsNull(rs.Fields.Item("Ruangan").Value), "", rs.Fields.Item("Ruangan").Value)
                    .SubItems(6) = IIf(IsNull(rs.Fields.Item("Jabatan").Value), "", rs.Fields.Item("Jabatan").Value)
                    .SubItems(7) = IIf(IsNull(rs.Fields.Item("Tgl. Daftar").Value), "", .SubItems(7) = rs.Fields.Item("Tgl. Daftar").Value)
                End With
            End If
        Else 'If lsv.ListItems.Count = jumlahTotalPIN Then
            For j = 1 To lsv.ListItems.Count
                If tempDataPIN(i) = lsv.ListItems.Item(j).Text Then
                    cr = InStr(1, lsv.ListItems(j).SubItems(1), frsTujuan, vbTextCompare)
                    If cr = 0 Then
                        '                        lsv.ListItems.Item(j).SubItems(1) = frsTujuan
                        '                    Else
                        lsv.ListItems.Item(j).SubItems(1) = lsv.ListItems.Item(j).SubItems(1) & "," & frsTujuan
                    End If
                    sama = sama + 1
                    GoTo lompat
                    '                    Exit For
                    '                Else
                    '                    GoTo terusin
                    '                    If rs.RecordCount = 0 Then
                    '                    'If tempDataPIN(i) = "" Then GoTo lanjutkan
                    '        '            frmPercobaan.lstPIN.AddItem tempDataPIN(i)
                    '                    lsv.ListItems.add(, , tempDataPIN(i)).SubItems(1) = frsTujuan
                    '                    unID = unID + 1
                    '        '            ls.SubItems(0) = tempDataPIN(i)
                    '                    Else
                    '                        With lsv.ListItems.add(, , tempDataPIN(i))
                    '                            .SubItems(1) = frsTujuan 'IIf(IsNull(rs.Fields.Item("Alamat FRS").Value), "", frsTujuan) 'rs.Fields.Item("Alamat FRS").Value)
                    '                            .SubItems(2) = IIf(IsNull(rs.Fields.Item("Nama").Value), "", rs.Fields.Item("Nama").Value)
                    '                            .SubItems(3) = IIf(IsNull(rs.Fields.Item("JK").Value), "", rs.Fields.Item("JK").Value)
                    '                            .SubItems(4) = IIf(IsNull(rs.Fields.Item("ID").Value), "", rs.Fields.Item("ID").Value)
                    '                            .SubItems(5) = IIf(IsNull(rs.Fields.Item("Ruangan").Value), "", rs.Fields.Item("Ruangan").Value)
                    '                            .SubItems(6) = IIf(IsNull(rs.Fields.Item("Jabatan").Value), "", rs.Fields.Item("Jabatan").Value)
                    '                            .SubItems(7) = IIf(IsNull(rs.Fields.Item("Tgl. Daftar").Value), "", .SubItems(7) = rs.Fields.Item("Tgl. Daftar").Value)
                    '                        End With
                    '                    End If
                End If
            Next
            If rs.RecordCount = 0 Then
                'If tempDataPIN(i) = "" Then GoTo lanjutkan
                '            frmPercobaan.lstPIN.AddItem tempDataPIN(i)
                lsv.ListItems.add(, , tempDataPIN(i)).SubItems(1) = frsTujuan
                unID = unID + 1
                '            ls.SubItems(0) = tempDataPIN(i)
            Else
                With lsv.ListItems.add(, , tempDataPIN(i))
                    If IsNull(rs.Fields.Item("Alamat FRS").Value) Then
                        .SubItems(1) = frsTujuan
                    Else
                        cr = InStr(1, rs.Fields.Item("Alamat FRS").Value, frsTujuan, vbTextCompare)
                        If cr = 0 Then
                            .SubItems(1) = rs.Fields.Item("Alamat FRS").Value & "," & frsTujuan
                        Else
                            .SubItems(1) = rs.Fields.Item("Alamat FRS").Value
                        End If
                    End If
                    '                    .SubItems(1) = frsTujuan   'IIf(IsNull(rs.Fields.Item("Alamat FRS").Value), "", frsTujuan) 'rs.Fields.Item("Alamat FRS").Value)
                    .SubItems(2) = IIf(IsNull(rs.Fields.Item("Nama").Value), "", rs.Fields.Item("Nama").Value)
                    .SubItems(3) = IIf(IsNull(rs.Fields.Item("JK").Value), "", rs.Fields.Item("JK").Value)
                    .SubItems(4) = IIf(IsNull(rs.Fields.Item("ID").Value), "", rs.Fields.Item("ID").Value)
                    .SubItems(5) = IIf(IsNull(rs.Fields.Item("Ruangan").Value), "", rs.Fields.Item("Ruangan").Value)
                    .SubItems(6) = IIf(IsNull(rs.Fields.Item("Jabatan").Value), "", rs.Fields.Item("Jabatan").Value)
                    .SubItems(7) = IIf(IsNull(rs.Fields.Item("Tgl. Daftar").Value), "", .SubItems(7) = rs.Fields.Item("Tgl. Daftar").Value)
                End With
            End If
lompat:
        End If
        N = N + 1
        pgb.Value = i
lanjutkan:
    Next
    '    lsv.ListItems = ls
    'frmPercobaan.lblJumlah.Caption = n & "; " & frmPercobaan.lstPIN.ListCount
    frsLihatPIN = frsTujuan
    If unID > 0 Then
        MsgBox "Jumlah PIN yang tidak memiliki data kepemilikan: " & unID & " dari " & lsv.ListItems.Count & " pin. Sama=" & sama, vbExclamation, "Perhatian"
    End If
End Sub

'=============== TAMBAHAN 6 Februari 2008 ================================
Public Function funcBuatProtokolUpload(ByVal frsUpload As String, ByVal pinUpload As String) As String
    Dim intFrsUpload As Integer
    Dim protokolUpload As String

    intFrsUpload = CInt(frsUpload)

    pinCek = ValidPIN(pinUpload)

    protokolUpload = Chr$(&H2) & Chr$(&HD) & Chr$(intFrsUpload) & Chr$(&H11) & _
    pinCek & Chr$(&H3)
    protokolUpload = protokolUpload & Chr$(cekSum(protokolUpload))

    funcBuatProtokolUpload = protokolUpload
End Function

'=============== TAMBAHAN 15 Februari 2008 ================================
Public Function funcConvertImageKeHex(ByVal imageData As Variant) As String
    Dim intLen As Integer, i As Integer
    Dim strHex As String, strHexAll As String
    Dim strImgData As String

    strImgData = CStr(imageData)
    intLen = Len(imageData)
    For i = 1 To intLen
        strHex = Hex$(Asc(Mid(imageData, i, 1)))
        If Len(strHex) = 1 Then strHex = "0" & strHex
        strHexAll = strHexAll & strHex
    Next
    funcConvertImageKeHex = strHexAll
End Function

Public Function funcConvertHexKeImage(ByVal strHexLine As String) As Variant
    Dim varImgData As Variant
    Dim intLen As Integer, i As Integer
    Dim strHex As String
    Dim intDes As Integer

    intLen = Len(strHexLine)
    For i = 1 To intLen Step 2
        strHex = Mid(strHexLine, i, 2)
        intDes = HexKeDesimal(strHex)
        varImgData = varImgData & Chr(intDes)
    Next
    funcConvertHexKeImage = varImgData
End Function

'=============== TAMBAHAN 21 Februari 2008 ================================
Public Sub subNetSend(ByVal hostTujuan As String, ByVal teksKirim As String)
    Shell "net send " & hostTujuan & " " & teksKirim, vbNormalFocus
End Sub

' add mathe 2009-06-04
Public Function funcErrorPrint(aErrorCode As Long) As String

    Select Case aErrorCode
        Case 0
            funcErrorPrint = "SUCCESS"
        Case 1
            funcErrorPrint = "ERR_COMPORT_ERROR"
        Case 2
            funcErrorPrint = "ERR_WRITE_FAIL"
        Case 3
            funcErrorPrint = "ERR_READ_FAIL"
        Case 4
            funcErrorPrint = "ERR_INVALID_PARAM"
        Case 5
            funcErrorPrint = "ERR_NON_CARRYOUT"
        Case 6
            funcErrorPrint = "ERR_LOG_END"
        Case 7
            funcErrorPrint = "ERR_MEMORY"
        Case 8
            funcErrorPrint = "ERR_MULTIUSER"
        Case 9
            funcErrorPrint = "ERR_NOSUPPORT"
    End Select
End Function


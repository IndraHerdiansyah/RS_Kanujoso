Attribute VB_Name = "ModuleAbsensi2"
Dim frs As String                   'frs = variabel jumlah alamat / alamat ke
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

If frs = add Then                       'apabila add(variabel alamat FRS yg dicari) = frs(variabel jumlah alat) maka :
    kirim_konfirmasi.Enabled = False    'mematikan timer kirim_konfirmasi
    a = 0                               'variabel baris msflexgrid(a) = 0 / direset
    frmKFRS.Show                        'menampilkan frmKFRS
Else
    frs = (frs + &H1)                   'apabila add <> frs maka alamat alat sekarang + 1
    MSComm1.Output = Chr$(&H2) & Chr$(&H5) & Chr$(frs) & Chr$(&H4) & Chr$(&H3) & Chr$(frs)  'kirim kode konfirmasi ke serial
End If

End Sub

'timer untuk mengirim kode untuk meminta hasil absensi
Private Sub minta_absensi_Timer()
  
If c = 1 Then                           'jika c (kode bantu tanda untuk meminta absen kembali) = 1

If frs = add2 Then                      'apabila frs = add2 maka :
    frs = &H0                           'frs direset ke alamat 0
End If

    code2 = code                        'isi code2 = code
    frs = (frs + &H1)                   'alamat alat sekarang + 1
    MSComm1.Output = Chr$(&H2) & Chr$(&H5) & Chr$(frs) & Chr$(&H1) & Chr$(&H3) & Chr$(code) 'kirim kode meminta absen ke serial
       
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
    MSComm1.Output = Chr$(&H6) & Chr$(&H2) & Chr$(&H5) & Chr$(frs) & Chr$(&H1) & Chr$(&H3) & Chr$(code2)    'kirim kode meminta absen kembali ke serial

End If

End Sub

Private Sub mnexit_Click()

Unload Me

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
    frs = &H0                           'frs direset
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


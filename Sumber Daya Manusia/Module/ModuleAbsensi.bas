Attribute VB_Name = "ModuleAbsensi"
Public i As Integer, Settings As String, Offset As Integer
Public add, add2 As String
Public inbuff1 As String
Public inbuff2 As Variant
Public a, c As String
Public fRS As String
Public code, code2 As String
Public indikator, indikator2, caristring, carichar, cek, simpanPIN, addpin As String
Public h1, h2, h3, h4, jumlah, hasil, R As String
Public frsHapus As Integer, pinHapus, h11, h12, h13, h14, h15, h16, h17, h18, pinCek, frsCek As String
Public key, fp As Integer, FP1, FP2, FP3, FP4, gmbrFP As String
Public pin1, pin2, pin3, pin4, pinz, cekz As String

Public frsTujuan As String, frsAsal As Integer, statDownload As Boolean  '9 Januari 2008
Public pinHapusFRS As String, pinHapusAsli As String, protokolHapusPin As String   '9 Januari 2008
Public statusPinHapus As Boolean, statusTransfer As Boolean  '9 Januari 2008

Public bolCekJumlahPIN As Boolean
Public bolPrepareFullUpload As Boolean
Public bolFullUpload As Boolean
Public protokolFullUpload As String
Public protokolPrepareFullUpload As String
Public jumlahPIN As Integer   'jumlah PIN pada satu FRS
Public jumlahTotalPIN As Integer
Public nPrepareUpload As Integer
Public tempDataPIN(1000) As String
Public idxTempDataPIN As Integer
Public statTransfer As Boolean
Public pinSimpan As String
Public alamatFRS As String
Public idxListViewPIN As Integer
Public frsLihatPIN As String
Public jumlahFRS As Integer
Public dl As Integer
Public varImageData As Variant
Public intBuatPIN As Integer
Public strStatusSekarang As String
Public resetInteger As Boolean

Public varFingerPrint() As typeFingerPrint
Public Type typeFingerPrint
    Lokasi As String
    IPAddress As String
    PortNumber As String
    Password As String
    Connected As Boolean
End Type

Public strTipeKoneksi As String
Public blnSettingsLoaded As Boolean

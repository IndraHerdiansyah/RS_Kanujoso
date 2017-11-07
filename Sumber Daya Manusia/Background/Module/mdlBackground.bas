Attribute VB_Name = "mdlBackground"
Option Explicit

Dim objRS As New ADODB.recordset
Dim DataFile As Integer, Fl As Long, Chunks As Integer
Dim Fragment As Integer, Chunk() As Byte, i As Integer, FileName As String

Private Const ChunkSize As Integer = 16384
Private Const conChunkSize = 100

Dim blnTrue As Boolean
Public strKdRS As String

Public Sub funShowBackground(ByVal lcNamaKota As String)
    
On Error GoTo errHand

    Dim frmBack As frmBackground
    Set frmBack = New frmBackground
    
    Dim strsql As String
    With objRS
        strsql = "SELECT * FROM LogoRumahSakit WHERE KdRS = '" & strKdRS & "'"
        .Open strsql, dbConn, adOpenForwardOnly, adLockReadOnly
            If Not .EOF Then
                blnTrue = True
                DataFile = 1
                Open "picTemp" For Binary Access Write As DataFile
                    Fl = objRS!Logo.ActualSize
                    If Fl = 0 Then Close DataFile: Exit Sub
                    Chunks = Fl \ ChunkSize
                    Fragment = Fl Mod ChunkSize
                    ReDim Chunk(Fragment)
                    Chunk() = objRS!Logo.GetChunk(Fragment)
                    Put DataFile, , Chunk()
                    For i = 1 To Chunks
                        ReDim Buffer(ChunkSize)
                        Chunk() = objRS!Logo.GetChunk(ChunkSize)
                        Put DataFile, , Chunk()
                    Next i
                Close DataFile
                
                FileName = "picTemp"
            Else
                blnTrue = False
            End If
        .Close
    End With
    
    Set objRS = Nothing

    With frmBack
        .lblNamaRS.Caption = "RSUD " & StrConv(lcNamaKota, vbProperCase)
        .imgLogoRS.Picture = LoadPicture(FileName)
        .Show
    End With
    
    Exit Sub

errHand:

    MsgBox Err.Number & vbCrLf & Err.Description
    
End Sub

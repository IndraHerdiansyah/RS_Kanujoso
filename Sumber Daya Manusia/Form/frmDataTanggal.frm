VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDataTanggal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Hari Libur"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmDataTanggal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   5295
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter Bulan"
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
      Left            =   120
      TabIndex        =   19
      Top             =   7320
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFilter 
      Height          =   375
      Left            =   1320
      TabIndex        =   18
      Top             =   7320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "MMMM yyyy"
      Format          =   113704963
      UpDown          =   -1  'True
      CurrentDate     =   39554
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Buat Tanggal"
      Height          =   315
      Left            =   3960
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpTgl 
      Height          =   315
      Left            =   720
      TabIndex        =   16
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "DD-MMM-YYYY"
      Format          =   113704960
      UpDown          =   -1  'True
      CurrentDate     =   39310
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
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
      Left            =   2400
      TabIndex        =   15
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
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
      Left            =   3375
      TabIndex        =   14
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
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
      Left            =   4320
      TabIndex        =   13
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
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
      Left            =   1440
      TabIndex        =   12
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtKdtgl 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtkode 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtisi 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dgDataTanggal 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin MSDataListLib.DataCombo dcharilibur 
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3625
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Kode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   10
      Top             =   1080
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hari Libur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3600
      Picture         =   "frmDataTanggal.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataTanggal.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataTanggal.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDataTanggal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim i As Integer

Private Sub chkFilter_Click()

    Me.dtpFilter.Value = Now
    If Me.chkFilter.Value = 1 Then
        Me.dtpFilter.Enabled = True
        Call loadDataGrid(True)
    Else
        Me.dtpFilter.Enabled = False
        Call loadDataGrid
    End If
End Sub

Private Sub cmdBatal_Click()
    Call subKosong
    Call subLoadDcSource
    Call subSetGrid
'    Me.dcharilibur.Visible = False'//yayang.agus 2014-08-14
    dcharilibur.Text = ""
End Sub

Private Sub cmdGenerate_Click()
    Dim intBulan As Integer
    Dim intHasilBagiBulan As Integer
    Dim intJumHari As Integer
    Dim intTahun As Integer
    Dim strTahun As String
    Dim intHasilBagiTahun As Integer
    Dim i As Integer

    Call subSetGrid
    Me.txtkdtgl.Text = ""

    intBulan = Month(Me.dtpTgl.Value)
    If intBulan <= 7 Then
        intHasilBagiBulan = intBulan Mod 2
        If intHasilBagiBulan = 1 Then
            intJumHari = 31
        Else
            intJumHari = 30
        End If
    Else
        intHasilBagiBulan = intBulan Mod 2
        If intHasilBagiBulan = 1 Then
            intJumHari = 30
        Else
            intJumHari = 31
        End If
    End If
    strTahun = CStr(Year(Me.dtpTgl.Value))
    intTahun = CInt(Right(strTahun, 2))
    intHasilBagiTahun = intTahun Mod 4
    If intBulan = 2 Then
        If intHasilBagiTahun = 0 Then
            intJumHari = 29
        Else
            intJumHari = 28
        End If
    End If
    For i = 1 To intJumHari
        Me.dtpTgl.Day = i
        With fgData
            .TextMatrix(.Rows - 1, 1) = txtkdtgl.Text
            .TextMatrix(.Rows - 1, 2) = Format(dtpTgl.Value, "DD/MM/yyyy")
            .Rows = .Rows + 1
        End With
    Next
    Me.dtpTgl.Day = 1
End Sub

Private Sub cmdHapus_Click()
'//yayang.agus 2014-08-14
On Error GoTo errSimpan
    Dim blnSave As Boolean
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Set dbcmd = New ADODB.Command
    blnSave = sp_simpan(txtkdtgl.Text, Format(dtpTgl.Value, "DD/MM/yyyy"), dcharilibur.BoundText, "D")   'hapus

    If blnSave = True Then
'        MsgBox "Hapus data berhasil.", vbInformation, "Informasi"
        Call cmdBatal_Click
        If Me.chkFilter.Value = 1 Then
            Call loadDataGrid(True)
        Else
            Call loadDataGrid
        End If
        dcharilibur.Text = ""
    End If
    Exit Sub
errSimpan:
    Set dbcmd = Nothing
    msubPesanError
'//
    
'    On Error GoTo xxx
'    strSQL = "DELETE DataTanggal WHERE KdTgl = '" & txtKdtgl.Text & "'"
'    Set adoComm = New ADODB.Command
'    With adoComm
'        .ActiveConnection = dbConn
'        .CommandText = strSQL
'        .CommandType = adCmdText
'        .Execute
'    End With
'    Call loadDataGrid
'    Call subKosong
'    Call cmdBatal_Click
'    dcharilibur.Text = ""
'    Exit Sub
'xxx:
'    MsgBox "Data Gagal Dihapus." & vbCrLf & Err.Description, _
'    vbCritical, "Error"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan
    Dim blnSave As Boolean
    If fgData.Rows = 2 Then Exit Sub
    Set dbcmd = New ADODB.Command
    For i = 1 To fgData.Rows - 2
        Me.dcharilibur.Text = Me.fgData.TextMatrix(i, 3)
        'blnSave = sp_simpan(fgData.TextMatrix(i, 2), Me.dcharilibur.BoundText, "U")
        blnSave = sp_simpan(fgData.TextMatrix(i, 1), fgData.TextMatrix(i, 2), fgData.TextMatrix(i, 3), "U") '//yayang.agus 2014-08-14
    Next i

    If blnSave = True Then
        'MsgBox "Penyimpanan Berhasil, Kode Tanggal : " & txtkdtgl.Text, vbInformation, "Informasi"
        MsgBox "Penyimpanan Berhasil.", vbInformation, "Informasi" '//yayang.agus 2014-08-14
        Call cmdBatal_Click
        If Me.chkFilter.Value = 1 Then
            Call loadDataGrid(True)
        Else
            Call loadDataGrid
        End If
        dcharilibur.Text = ""
    End If
    Exit Sub
errSimpan:
    Set dbcmd = Nothing
    msubPesanError
End Sub

Private Sub cmdPlus_Click()
    On Error GoTo errLoad
    Dim i As Integer

    With fgData
        .TextMatrix(.Rows - 1, 1) = txtkdtgl.Text
        .TextMatrix(.Rows - 1, 2) = Format(dtpTgl.Value, "DD/MM/yyyy")
        .TextMatrix(.Rows - 1, 3) = dcharilibur.BoundText
        .TextMatrix(.Rows - 1, 4) = dcharilibur.Text
        .Rows = .Rows + 1
    End With

    txtkdtgl.Text = ""
'    dtpTgl.Value = Now'//yayang.agus 2014-08-14

    dcharilibur.Text = ""
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdmin_Click()
    On Error GoTo errLoad
    Dim i As Integer
    With fgData
        If .row = .Rows Then Exit Sub
        If .row = 0 Then Exit Sub

        If .Rows = 2 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next i
            Exit Sub
        Else
            .RemoveItem .row
        End If
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcharilibur_Change()
'//yayang.agus 2014-08-14
'    Me.fgData.TextMatrix(Me.fgData.row, Me.fgData.Col) = Me.dcharilibur.Text
'    Me.dcharilibur.Visible = False
'//
    Call dcharilibur_KeyPress(13)
End Sub

Private Sub dcharilibur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
'        Me.dcharilibur.Visible = False'//yayang.agus 2014-08-14
    End If
End Sub

Private Sub dcharilibur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        Me.fgData.TextMatrix(Me.fgData.row, Me.fgData.Col) = Me.dcharilibur.Text'//yayang.agus 2014-08-14
'        Me.dcharilibur.Visible = False'//yayang.agus 2014-08-14
    End If
End Sub

Private Sub dgDataTanggal_Click()
    On Error Resume Next
'    Call subSetGrid'//yayang.agus 2014-08-14
    With dgDataTanggal
        txtkdtgl.Text = .Columns(0).Text
'        Me.fgData.TextMatrix(1, 1) = .Columns(0).Value'//yayang.agus 2014-08-14
        
'        Me.fgData.TextMatrix(1, 2) = .Columns(1).Value'//yayang.agus 2014-08-14
        If .Columns(2).Text = "" Then
            dcharilibur.Text = ""
            dtpTgl.Value = Date
'            Me.fgData.TextMatrix(1, 3) = ""'//yayang.agus 2014-08-14
        Else
            dcharilibur.Text = .Columns(2).Text
            dtpTgl.Value = .Columns(1).Text
'            Me.fgData.TextMatrix(1, 3) = .Columns(2).Value'//yayang.agus 2014-08-14
        End If
'        Me.fgData.Rows = Me.fgData.Rows + 1'//yayang.agus 2014-08-14
    End With
    Exit Sub
End Sub

Private Sub dgDataTanggal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
dgDataTanggal_Click
End Sub

Private Sub dtpFilter_Change()
    Call loadDataGrid(True)
End Sub


Private Sub dtpTgl_Change()
'    Call loadDataGrid(False)'//yayang.agus 2014-08-14
End Sub

Private Sub dtpTgl_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then cmdGenerate.SetFocus'//yayang.agus 2014-08-14
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyPress(13)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If Me.fgData.Col <> 3 Then Exit Sub
        Me.fgData.TextMatrix(Me.fgData.row, Me.fgData.Col) = ""
    End If
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer) '//yayang.agus 2014-08-14
'    If KeyAscii = 13 Then
'        If Me.fgData.Col <> 3 Then Exit Sub
'        If Me.fgData.TextMatrix(Me.fgData.row, Me.fgData.Col - 1) = "" Then Exit Sub
'
'        Me.dcharilibur.Text = Me.fgData.TextMatrix(Me.fgData.row, Me.fgData.Col)
'        Me.dcharilibur.Left = Me.fgData.Left
'        Me.dcharilibur.Top = Me.fgData.Top
'
'        For i = 0 To fgData.Col - 1
'            dcharilibur.Left = dcharilibur.Left + fgData.ColWidth(i)
'        Next i
'        For i = 0 To fgData.row - 1
'            dcharilibur.Top = dcharilibur.Top + fgData.RowHeight(i)
'        Next i
'        If fgData.TopRow > 1 Then
'            dcharilibur.Top = dcharilibur.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
'        End If
'        Me.dcharilibur.Width = fgData.ColWidth(fgData.Col)
'        Me.dcharilibur.Visible = True
'        Me.dcharilibur.SetFocus
'        KeyAscii = 0
'    End If
End Sub

Private Sub Form_Activate()
    dtpTgl.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTgl.Value = Now

    Call subSetGrid
    Call subLoadDcSource
    Call loadDataGrid
    Call subKosong

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subKosong()
    txtkdtgl.Text = ""
    dtpTgl.Value = Now
    dcharilibur.Text = ""
End Sub

Private Sub subSetGrid()
    On Error GoTo errLoad
    With fgData
        .clear
        .Rows = 2
        .Cols = 5

        .RowHeight(0) = 500

        .TextMatrix(0, 1) = "Kode"
        .TextMatrix(0, 2) = "Tanggal"
        .TextMatrix(0, 4) = "Hari Libur"

        .ColWidth(0) = 0
        .ColWidth(1) = 500
        .ColWidth(2) = 1500
        .ColWidth(3) = 0
        .ColWidth(4) = 1500
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcharilibur, rs, "SELECT KdHariLibur, NamaHariLibur FROM HariLibur ORDER BY NamaHariLibur")
    If rs.EOF = False Then dcharilibur.BoundText = rs(0).Value

    Exit Sub
errLoad:
    Call msubPesanError
End Sub
'//yayang.agus 2014-08-14
'Private Function sp_simpan( f_NamaTgl As Date, f_KdHariLibur As String, f_status As String) As Boolean
Private Function sp_simpan(KdTgl As String, f_NamaTgl As Date, f_KdHariLibur As String, f_status As String) As Boolean
    On Error GoTo errLoad
    sp_simpan = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        '.Parameters.Append .CreateParameter("KdTgl", adVarChar, adParamInput, 3, txtKdtgl.Text)
        .Parameters.Append .CreateParameter("KdTgl", adVarChar, adParamInput, 3, KdTgl) '//yayang.agus 2014-08-14
        .Parameters.Append .CreateParameter("NamaTgl", adDate, adParamInput, , Format(f_NamaTgl, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("KdHariLibur", adVarChar, adParamInput, 3, IIf(f_KdHariLibur = "", Null, f_KdHariLibur))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        .Parameters.Append .CreateParameter("OutputKdTgl", adVarChar, adParamOutput, 3, Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_DataTanggal"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_simpan = False
        Else
            fgData.TextMatrix(i, 1) = IIf(IsNull(.Parameters("OutputKdTgl").Value), "", .Parameters("OutputKdTgl").Value)
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
errLoad:
    Call msubPesanError(" sp_Simpan")
    sp_simpan = False
End Function

Sub loadDataGrid(Optional ByVal blnFilter As Boolean)
    Set rs = Nothing
    If blnFilter Then
        strSQL = "select * from v_tanggal" & _
        " WHERE MONTH(Tanggal)='" & Month(Me.dtpFilter.Value) & "'" & _
        " AND YEAR(Tanggal)='" & Year(Me.dtpFilter.Value) & "' AND [Hari Libur]<>''"
    Else
'        strSQL = "select * from v_tanggal WHERE [Hari Libur]<>''"
        'strSQL = "select * from v_tanggal"
        '//yayang.agus 2014-08-08
        strSQL = "select * from v_tanggal" & _
        " WHERE MONTH(Tanggal)='" & Month(Me.dtpTgl.Value) & "'" & _
        " AND YEAR(Tanggal)='" & Year(Me.dtpTgl.Value) & "' " 'AND [Hari Libur]<>''"
        '//
    End If
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgDataTanggal.DataSource = rs
    With dgDataTanggal
        .Columns(0).Caption = "Kode"
        .Columns(1).Caption = "Tanggal"
        .Columns(2).Caption = "Hari Libur"
    End With
End Sub



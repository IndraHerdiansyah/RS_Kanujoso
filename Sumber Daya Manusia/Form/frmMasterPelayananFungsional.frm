VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterPelayananFungsional 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Pelayanan Fungsional"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterPelayananFungsional.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10680
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   8805
      TabIndex        =   12
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   7110
      TabIndex        =   11
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   5415
      TabIndex        =   10
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   6765
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   10620
      Begin TabDlg.SSTab SSTab1 
         Height          =   6330
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   11165
         _Version        =   393216
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Jenis Pelayanan Fungsional"
         TabPicture(0)   =   "frmMasterPelayananFungsional.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Daftar Pelayanan Fungsional"
         TabPicture(1)   =   "frmMasterPelayananFungsional.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Nilai Kredit "
         TabPicture(2)   =   "frmMasterPelayananFungsional.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame4"
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame4 
            Height          =   5655
            Left            =   -74760
            TabIndex        =   16
            Top             =   480
            Width           =   9915
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   330
               Left            =   5520
               MaxLength       =   50
               TabIndex        =   34
               Top             =   360
               Width           =   4200
            End
            Begin VB.TextBox txtNilai 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   6600
               MaxLength       =   5
               TabIndex        =   31
               Top             =   840
               Width           =   1560
            End
            Begin MSDataListLib.DataCombo dcNJP 
               Height          =   330
               Left            =   1680
               TabIndex        =   8
               Top             =   360
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataListLib.DataCombo dcJabatanF 
               Height          =   330
               Left            =   1680
               TabIndex        =   29
               Top             =   840
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataGridLib.DataGrid dgNilai 
               Height          =   4065
               Left            =   240
               TabIndex        =   33
               Top             =   1320
               Width           =   9480
               _ExtentX        =   16722
               _ExtentY        =   7170
               _Version        =   393216
               AllowUpdate     =   -1  'True
               HeadLines       =   2
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
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
                  AllowRowSizing  =   0   'False
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Nilai Kredit"
               Height          =   210
               Left            =   5520
               TabIndex        =   32
               Top             =   840
               Width           =   840
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Jabatan Fungsi"
               Height          =   210
               Left            =   240
               TabIndex        =   30
               Top             =   840
               Width           =   1200
            End
            Begin VB.Label lblJmlDataDetailPelayanan 
               AutoSize        =   -1  'True
               Caption         =   "Data Ke / "
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   240
               TabIndex        =   27
               Top             =   5640
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Nama Pelayanan"
               Height          =   210
               Left            =   240
               TabIndex        =   22
               Top             =   360
               Width           =   1320
            End
         End
         Begin VB.Frame Frame3 
            Height          =   5655
            Left            =   -74760
            TabIndex        =   15
            Top             =   480
            Width           =   9915
            Begin VB.TextBox txtCariJnsPelayanan 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2040
               TabIndex        =   25
               Top             =   5160
               Width           =   4815
            End
            Begin VB.TextBox txtKDP 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   240
               MaxLength       =   6
               TabIndex        =   4
               Top             =   600
               Width           =   840
            End
            Begin VB.TextBox txtNDP 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   4080
               MaxLength       =   75
               TabIndex        =   5
               Top             =   600
               Width           =   5640
            End
            Begin MSDataGridLib.DataGrid dgDaftarPelayanan 
               Height          =   3945
               Left            =   240
               TabIndex        =   7
               Top             =   1080
               Width           =   9480
               _ExtentX        =   16722
               _ExtentY        =   6959
               _Version        =   393216
               AllowUpdate     =   -1  'True
               HeadLines       =   2
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
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
                  AllowRowSizing  =   0   'False
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin MSDataListLib.DataCombo dcJnsPelayanan 
               Height          =   330
               Left            =   1200
               TabIndex        =   6
               Top             =   600
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin VB.Label Label7 
               Caption         =   "Cari Nama Pelayanan"
               Height          =   255
               Left            =   240
               TabIndex        =   26
               Top             =   5160
               Width           =   2415
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Jenis Pelayanan"
               Height          =   210
               Left            =   1200
               TabIndex        =   21
               Top             =   360
               Width           =   1260
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Kode"
               Height          =   210
               Left            =   240
               TabIndex        =   20
               Top             =   360
               Width           =   420
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Nama Pelayanan"
               Height          =   210
               Left            =   4080
               TabIndex        =   19
               Top             =   360
               Width           =   1320
            End
         End
         Begin VB.Frame Frame2 
            Height          =   5655
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   9915
            Begin VB.TextBox txtIND 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               TabIndex        =   23
               Top             =   5160
               Width           =   4815
            End
            Begin VB.TextBox txtNJP 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1200
               MaxLength       =   50
               TabIndex        =   2
               Top             =   600
               Width           =   6480
            End
            Begin VB.TextBox txtKJP 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   240
               MaxLength       =   3
               TabIndex        =   1
               Top             =   600
               Width           =   840
            End
            Begin MSDataGridLib.DataGrid dgJenisPelayanan 
               Height          =   3975
               Left            =   240
               TabIndex        =   3
               Top             =   1080
               Width           =   9480
               _ExtentX        =   16722
               _ExtentY        =   7011
               _Version        =   393216
               AllowUpdate     =   0   'False
               HeadLines       =   2
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
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
                  AllowRowSizing  =   0   'False
                  Locked          =   -1  'True
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label1 
               Caption         =   "Cari Jenis Pelayanan"
               Height          =   255
               Left            =   240
               TabIndex        =   24
               Top             =   5205
               Width           =   2415
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Jenis Pelayanan"
               Height          =   210
               Left            =   1200
               TabIndex        =   18
               Top             =   360
               Width           =   1770
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Kode"
               Height          =   210
               Left            =   240
               TabIndex        =   17
               Top             =   360
               Width           =   420
            End
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   28
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterPelayananFungsional.frx":0D1E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8880
      Picture         =   "frmMasterPelayananFungsional.frx":36DF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterPelayananFungsional.frx":4467
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMasterPelayananFungsional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    On Error GoTo errLoad

    Call subKosong
    Call subLoadGridSource
    Call subDcSource
    Call SSTab1_KeyPress(13)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdDel_Click()
    On Error GoTo errLoad

    Select Case SSTab1.Tab

        Case 0
            If Periksa("text", txtNJP, "Silahkan isi Nama Jenis Pelayanan") = False Then Exit Sub

            If sp_JenisPelayanan("D") = False Then Exit Sub

        Case 1
            If Periksa("datacombo", dcJnsPelayanan, "Silahkan isi Jenis Pelayanan Fungsional") = False Then Exit Sub
            If Periksa("text", txtNDP, "Nama Pelayanan Fungsional kosong ") = False Then Exit Sub

            If sp_ListPelayanan("D") = False Then Exit Sub

        Case 2
            If Periksa("datacombo", dcNJP, "Nama Pelayanan kosong") = False Then Exit Sub
            If Periksa("datacombo", dcJabatanF, "Jabatan fungsi kosong") = False Then Exit Sub
            If Periksa("text", txtNilai, "Silahkan isi Nilai") = False Then Exit Sub

            If sp_Nilai("D") = False Then Exit Sub

    End Select

    MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
    Call cmdCancel_Click

    Exit Sub
errLoad:
    MsgBox "Penghapusan data gagal", vbCritical, "Validasi"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errLoad

    Select Case SSTab1.Tab
        Case 0

            If Periksa("text", txtNJP, "Silahkan isi Nama Jenis Pelayanan") = False Then Exit Sub

            If sp_JenisPelayanan("A") = False Then Exit Sub
            Call cmdCancel_Click
        Case 1
            If Periksa("datacombo", dcJnsPelayanan, "Silahkan isi Jenis Pelayanan Fungsional") = False Then Exit Sub
            If Periksa("text", txtNDP, "Nama Pelayanan Fungsional kosong ") = False Then Exit Sub

            If sp_ListPelayanan("A") = False Then Exit Sub
            Call cmdCancel_Click

        Case 2
            If Periksa("datacombo", dcNJP, "Nama Pelayanan kosong") = False Then Exit Sub
            If Periksa("datacombo", dcJabatanF, "Jabatan fungsi kosong") = False Then Exit Sub
            If Periksa("text", txtNilai, "Silahkan isi Nilai") = False Then Exit Sub

            If sp_Nilai("A") = False Then Exit Sub
            Call cmdCancel_Click

    End Select

    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJabatanF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNilai.SetFocus
End Sub

Private Sub dcJnsPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNDP.SetFocus
End Sub

Private Sub dcNJP_Change()
    On Error GoTo errLoad
    If dcNJP.MatchedWithList = False Then Text1.Text = "": Exit Sub
    strSQL = "Select ListPelayananFungsional.KdPelayananF, ListPelayananFungsional.NamaPelayananF, JenisPelayananFungsional.JenisPelayananF" & _
    " FROM ListPelayananFungsional INNER JOIN JenisPelayananFungsional ON ListPelayananFungsional.KdJenisPelayananF = JenisPelayananFungsional.KdJenisPelayananF" & _
    " WHERE NamaPelayananF LIKE '" & dcNJP.Text & "'" & _
    " ORDER BY NamaPelayananF"
    Call msubRecFO(rsb, strSQL)
    If rsb.EOF = True Then Exit Sub
    dcNJP.BoundText = rsb(0).Value: Text1.Text = rsb("JenisPelayananF").Value

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNJP_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        If Len(Trim(dcNJP.Text)) = 0 Then Exit Sub
        If Text1.Text <> "" Then GoTo stepNext
        strSQL = "Select ListPelayananFungsional.KdPelayananF, ListPelayananFungsional.NamaPelayananF, JenisPelayananFungsioanl.JenisPelayananF" & _
        " FROM ListPelayananFungsional INNER JOIN JenisPelayananFungsional ON ListPelayananFungsional.KdJenisPelayananF = JenisPelayananFungsional.KdJenisPelayananF" & _
        " WHERE NamaPelayananF LIKE '%" & dcNJP.Text & "%'" & _
        " ORDER BY NamaPelayananF"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Text1.Text = "": Exit Sub
        dcNJP.BoundText = rs(0).Value: Text1.Text = rs("JenisPelayananF").Value
stepNext:
        dcJabatanF.SetFocus
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgDaftarPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNDP.SetFocus
End Sub

Private Sub dgDaftarPelayanan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad
    If dgDaftarPelayanan.ApproxCount = 0 Then Exit Sub
    txtKDP.Text = dgDaftarPelayanan.Columns(0).Value
    txtNDP.Text = dgDaftarPelayanan.Columns(1).Value
    dcJnsPelayanan.Text = dgDaftarPelayanan.Columns(2).Value
    Exit Sub
errLoad:
    dcJnsPelayanan.BoundText = ""
End Sub

Private Sub dgJenisPelayanan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad
    If dgJenisPelayanan.ApproxCount = 0 Then Exit Sub
    txtKJP = dgJenisPelayanan.Columns(0).Value
    txtNJP = dgJenisPelayanan.Columns(1).Value
    Exit Sub
errLoad:
End Sub

Private Sub dgNilai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad
    If dgNilai.ApproxCount = 0 Then Exit Sub
    dcNJP.BoundText = dgNilai.Columns(0).Value
    Text1.Text = dgNilai.Columns(6).Value
    dcJabatanF.BoundText = dgNilai.Columns(2).Value
    txtNilai.Text = dgNilai.Columns(4).Value
    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKey1
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 0
        Case vbKey2
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 1
        Case vbKey3
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 2

    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    SSTab1.Tab = 0
    Call cmdCancel_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub lvDetailPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub lvPelayananRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error GoTo errLoad

    Call cmdCancel_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub subKosong()
    Select Case SSTab1.Tab
        Case 0
            txtKJP = ""
            txtNJP = ""

        Case 1
            txtKDP = ""
            txtNDP = ""
            dcJnsPelayanan.BoundText = ""

        Case 2
            dcNJP.BoundText = ""
            dcJabatanF.BoundText = ""
            Text1.Text = ""
            txtNilai.Text = ""
    End Select
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad

    Select Case SSTab1.Tab
        Case 0
            strSQL = "Select * From JenisPelayananFungsional ORDER BY JenisPelayananF"
            Call msubRecFO(rs, strSQL)
            Set dgJenisPelayanan.DataSource = rs
            Call setJenisPelayanan

        Case 1
            strSQL = "SELECT dbo.ListPelayananFungsional.KdPelayananF, dbo.ListPelayananFungsional.NamaPelayananF, dbo.JenisPelayananFungsional.JenisPelayananF FROM dbo.ListPelayananFungsional INNER JOIN dbo.JenisPelayananFungsional ON dbo.ListPelayananFungsional.KdJenisPelayananF = dbo.JenisPelayananFungsional.KdJenisPelayananF"
            Call msubRecFO(rs, strSQL)
            Set dgDaftarPelayanan.DataSource = rs
            Call setDaftarPelayanan

        Case 2
            strSQL = "SELECT dbo.NilaiIndexPelayanan.KdPelayananF, dbo.ListPelayananFungsional.NamaPelayananF, dbo.NilaiIndexPelayanan.KdDetailJabatanF, " & _
            "dbo.DetailJabatanFungsional.DetailJabatanF, dbo.NilaiIndexPelayanan.Nilai, dbo.JenisPelayananFungsional.KdJenisPelayananF, " & _
            "dbo.JenisPelayananFungsional.JenisPelayananF " & _
            "FROM dbo.NilaiIndexPelayanan INNER JOIN " & _
            "dbo.ListPelayananFungsional ON dbo.NilaiIndexPelayanan.KdPelayananF = dbo.ListPelayananFungsional.KdPelayananF INNER JOIN " & _
            "dbo.DetailJabatanFungsional ON dbo.NilaiIndexPelayanan.KdDetailJabatanF = dbo.DetailJabatanFungsional.KdDetailJabatanF INNER JOIN " & _
            "dbo.JenisPelayananFungsional ON dbo.ListPelayananFungsional.KdJenisPelayananF = dbo.JenisPelayananFungsional.KdJenisPelayananF "
            Call msubRecFO(rs, strSQL)
            Set dgNilai.DataSource = rs
            Call setNilai
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_JenisPelayanan(f_Status As String) As Boolean
    sp_JenisPelayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisPelayananF", adChar, adParamInput, 3, txtKJP.Text)
        .Parameters.Append .CreateParameter("JenisPelayananF", adVarChar, adParamInput, 50, Trim(txtNJP.Text))
        .Parameters.Append .CreateParameter("OutputKdJnsPelayanan", adChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("StatusCode", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_JenisPelayananFungsional"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data", vbExclamation, "Validasi"
        Else
            If Not IsNull(.Parameters("OutputKdJnsPelayanan").Value) Then txtKJP = .Parameters("OutputKdJnsPelayanan").Value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_ListPelayanan(f_Status As String) As Boolean
    sp_ListPelayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisPelayananF", adChar, adParamInput, 3, dcJnsPelayanan.BoundText)
        .Parameters.Append .CreateParameter("NamaPelayananF", adVarChar, adParamInput, 75, txtNDP.Text)
        .Parameters.Append .CreateParameter("OutputKdPelayananRS", adChar, adParamInputOutput, 6, txtKDP.Text)
        .Parameters.Append .CreateParameter("StatusCode", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_ListPelayananFungsional"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data", vbCritical, "Validasi"
        Else
            If Not IsNull(.Parameters("OutputKdPelayananRS").Value) Then txtKDP.Text = .Parameters("OutputKdPelayananRS").Value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_Nilai(f_Status As String) As Boolean
    sp_Nilai = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPelayananF", adChar, adParamInput, 6, dcNJP.BoundText)
        .Parameters.Append .CreateParameter("KdDetailJabatanF", adChar, adParamInput, 5, dcJabatanF.BoundText)
        .Parameters.Append .CreateParameter("Nilai", adVarChar, adParamInput, 5, Trim(txtNilai.Text))
        .Parameters.Append .CreateParameter("StatusCode", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_NilaiPelayanan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data", vbExclamation, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Sub setJenisPelayanan()
    With dgJenisPelayanan
        .Columns(0).Width = 700
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 7200
        .Columns(0).Caption = "Kode"
        .Columns(1).Caption = "Nama Jenis Pelayanan"
    End With
End Sub

Sub setDaftarPelayanan()
    With dgDaftarPelayanan
        .Columns(0).Width = 1000
        .Columns(1).Width = 4380
        .Columns(2).Width = 3500

        .Columns(0).Caption = "Kode"
        .Columns(1).Caption = "Nama Pelayanan"
        .Columns(2).Caption = "Nama Jenis Pelayanan"
    End With
End Sub

Sub setNilai()
    With dgNilai
        .Columns(0).Width = 1000
        .Columns(1).Width = 4380
        .Columns(2).Width = 0
        .Columns(3).Width = 2000
        .Columns(4).Width = 2000
        .Columns(5).Width = 0
        .Columns(6).Width = 2000

        .Columns(0).Caption = "Kode"
        .Columns(1).Caption = "Nama Pelayanan"
        .Columns(3).Caption = "Jabatan Fungsional"
        .Columns(4).Caption = "Nilai Kredit"
        .Columns(6).Caption = "Jenis Pelayanan Fungsi"
    End With
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    If KeyAscii = 13 Then
        Select Case SSTab1.Tab
            Case 0
                txtNJP.SetFocus
            Case 1
                dcJnsPelayanan.SetFocus
            Case 2
                dcNJP.SetFocus

        End Select
    End If
errLoad:
End Sub

Private Sub txtCariJnsPelayanan_Change()
    On Error GoTo errLoad

    strSQL = "SELECT dbo.ListPelayananFungsional.KdPelayananF, dbo.ListPelayananFungsional.NamaPelayananF, dbo.JenisPelayananFungsional.JenisPelayananF FROM dbo.ListPelayananFungsional INNER JOIN dbo.JenisPelayananFungsional ON dbo.ListPelayananFungsional.KdJenisPelayananF = dbo.JenisPelayananFungsional.KdJenisPelayananF And dbo.ListPelayananFungsional.NamaPelayananF like '%" & txtCariJnsPelayanan.Text & "%'"
    Call msubRecFO(rs, strSQL)
    Set dgDaftarPelayanan.DataSource = rs
    Call setDaftarPelayanan

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIND_Change()
    On Error GoTo errLoad

    strSQL = "SELECT * FROM jenispelayananfungsional WHERE JenisPelayananF like'%" & txtIND & "%' ORDER BY JenisPelayananF"
    Call msubRecFO(rs, strSQL)
    Set dgJenisPelayanan.DataSource = rs
    Call setJenisPelayanan

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIND_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dgJenisPelayanan.SetFocus
End Sub

Private Sub subDcSource()
    On Error GoTo errLoad

    Select Case SSTab1.Tab

        Case 1
            strSQL = "Select * From JenisPelayananFungsional order by JenisPelayananF"
            Call msubDcSource(dcJnsPelayanan, rs, strSQL)

        Case 2
            strSQL = "SELECT * FROM ListPelayananFungsional order by NamaPelayananF"
            Call msubDcSource(dcNJP, rs, strSQL)

            strSQL = "SELECT * FROM DetailJabatanFungsional order by DetailJabatanF"
            Call msubDcSource(dcJabatanF, rs, strSQL)

    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNJP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtNDP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub


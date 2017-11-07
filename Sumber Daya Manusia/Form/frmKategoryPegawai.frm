VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmKategoryPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kategory & Detail Kategory Pegawai"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKategoryPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   2160
      TabIndex        =   33
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3480
      TabIndex        =   29
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4800
      TabIndex        =   30
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6120
      TabIndex        =   31
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   7440
      TabIndex        =   32
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6930
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   8745
      Begin TabDlg.SSTab SSTab1 
         Height          =   6570
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   11589
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Kategory Pegawai"
         TabPicture(0)   =   "frmKategoryPegawai.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail Kategory Pegawai"
         TabPicture(1)   =   "frmKategoryPegawai.frx":0CE6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6015
            Left            =   -74880
            TabIndex        =   18
            Top             =   480
            Width           =   8175
            Begin VB.TextBox txtParameter 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   4920
               MaxLength       =   30
               TabIndex        =   36
               Top             =   5570
               Width           =   3135
            End
            Begin VB.CheckBox CheckStatusEnbl 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               Height          =   255
               Left            =   5880
               TabIndex        =   34
               Top             =   1440
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtNamaExt 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   30
               TabIndex        =   5
               Top             =   1800
               Width           =   5055
            End
            Begin VB.TextBox txtKdExt 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   15
               TabIndex        =   4
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox txtRepDisplay 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   30
               TabIndex        =   3
               Top             =   1080
               Width           =   5055
            End
            Begin VB.TextBox txtKdKategory 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   2160
               MaxLength       =   1
               TabIndex        =   1
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtKategory 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   30
               TabIndex        =   2
               Top             =   720
               Width           =   5055
            End
            Begin MSDataGridLib.DataGrid dgKategory 
               Height          =   3195
               Left            =   120
               TabIndex        =   6
               Top             =   2280
               Width           =   7920
               _ExtentX        =   13970
               _ExtentY        =   5636
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
               Caption         =   "Cari Kategory Pegawai"
               Height          =   210
               Left            =   2880
               TabIndex        =   37
               Top             =   5640
               Width           =   1815
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nama external"
               Height          =   210
               Left            =   480
               TabIndex        =   25
               Top             =   1800
               Width           =   1170
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   480
               TabIndex        =   24
               Top             =   1440
               Width           =   1140
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Report Display"
               Height          =   210
               Left            =   480
               TabIndex        =   23
               Top             =   1080
               Width           =   1155
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Kategory Pegawai"
               Height          =   210
               Left            =   480
               TabIndex        =   20
               Top             =   735
               Width           =   1470
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Kode Kategory"
               Height          =   210
               Left            =   480
               TabIndex        =   19
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5970
            Left            =   120
            TabIndex        =   15
            Top             =   525
            Width           =   8265
            Begin VB.TextBox txtParameterDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   4680
               MaxLength       =   50
               TabIndex        =   38
               Top             =   5520
               Width           =   3480
            End
            Begin VB.CheckBox CheckStatusEnbl1 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               Height          =   255
               Left            =   5880
               TabIndex        =   35
               Top             =   1680
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtNamaExtDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   30
               TabIndex        =   12
               Top             =   2040
               Width           =   5040
            End
            Begin VB.TextBox txtKdExtDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   15
               TabIndex        =   11
               Top             =   1680
               Width           =   1815
            End
            Begin VB.TextBox txtRepDisplayDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   10
               Top             =   1320
               Width           =   5040
            End
            Begin VB.TextBox txtNamaDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   9
               Top             =   960
               Width           =   5040
            End
            Begin VB.TextBox txtKdDetail 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   2160
               MaxLength       =   2
               TabIndex        =   7
               Top             =   240
               Width           =   840
            End
            Begin MSDataGridLib.DataGrid dgDetailKategory 
               Height          =   2895
               Left            =   120
               TabIndex        =   13
               Top             =   2520
               Width           =   8055
               _ExtentX        =   14208
               _ExtentY        =   5106
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
            Begin MSDataListLib.DataCombo dcKategory 
               Height          =   330
               Left            =   2160
               TabIndex        =   8
               Top             =   600
               Width           =   3120
               _ExtentX        =   5503
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
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Cari Nama Detail"
               Height          =   210
               Left            =   3000
               TabIndex        =   39
               Top             =   5595
               Width           =   1545
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Nama External"
               Height          =   210
               Left            =   480
               TabIndex        =   28
               Top             =   2100
               Width           =   1170
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   480
               TabIndex        =   27
               Top             =   1740
               Width           =   1140
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Report Display"
               Height          =   210
               Left            =   480
               TabIndex        =   26
               Top             =   1380
               Width           =   1155
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Kategory Pegawai"
               Height          =   210
               Left            =   480
               TabIndex        =   21
               Top             =   645
               Width           =   1470
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Detail"
               Height          =   210
               Left            =   480
               TabIndex        =   17
               Top             =   1020
               Width           =   960
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Kode Detail"
               Height          =   210
               Left            =   480
               TabIndex        =   16
               Top             =   285
               Width           =   930
            End
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
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
   Begin VB.Image Image4 
      Height          =   945
      Left            =   7080
      Picture         =   "frmKategoryPegawai.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKategoryPegawai.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmKategoryPegawai.frx":30E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmKategoryPegawai.frx":5AA9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmKategoryPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub blankfield()
    On Error Resume Next
    Select Case SSTab1.Tab
        Case 0
            txtKdKategory.Text = ""
            txtKategory.Text = ""
            txtRepDisplay.Text = ""
            txtKdExt.Text = ""
            txtNamaExt.Text = ""
            txtKategory.SetFocus
            CheckStatusEnbl.Value = 1
        Case 1
            dcKategory = ""
            txtKdDetail = ""
            txtNamaDetail = ""
            txtRepDisplayDetail.Text = ""
            txtKdExtDetail.Text = ""
            txtNamaExtDetail.Text = ""
            txtNamaDetail.SetFocus
            CheckStatusEnbl1.Value = 1
    End Select
End Sub

Sub Dag()
    On Error GoTo errLoad
    Select Case SSTab1.Tab
        Case 0
            strSQL = "SELECT KdKategoryPegawai AS Kode, KategoryPegawai AS [Nama Kategory], ReportDisplay AS [Report Display], KodeExternal AS [Kd.Ext], NamaExternal AS [Nama Ext], StatusEnabled AS Status FROM KategoryPegawai where KategoryPegawai LIKE '%" & txtParameter.Text & "%'" 'WHERE (StatusEnabled <> 0) OR (StatusEnabled IS NULL)"
            Call msubRecFO(rs, strSQL)
            Set dgKategory.DataSource = rs
            dgKategory.Columns(0).Width = 500
            dgKategory.Columns(1).Width = 3000
            dgKategory.Columns(5).Width = 1000
        Case 1
            strSQL = "SELECT * from V_KategoryPegawai where [Nama Detail] LIKE '%" & txtParameterDetail.Text & "%' "
            Call msubRecFO(rs, strSQL)
            Set dgDetailKategory.DataSource = rs
            dgDetailKategory.Columns(1).Width = 2500
            dgDetailKategory.Columns(5).Width = 1000
            dgDetailKategory.Columns(7).Width = 0
            Call msubDcSource(dcKategory, rs, "SELECT KdKategoryPegawai, KategoryPegawai FROM KategoryPegawai where statusenabled='1'")
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Function sp_KategoryPegawai(f_status As String) As Boolean
    sp_KategoryPegawai = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKategoryPegawai", adChar, adParamInput, 1, txtKdKategory.Text)
        .Parameters.Append .CreateParameter("KategoryPegawai", adVarChar, adParamInput, 30, Trim(txtKategory.Text))
        .Parameters.Append .CreateParameter("ReportDisplay", adVarChar, adParamInput, 30, IIf(txtRepDisplay.Text = "", Null, Trim(txtKategory.Text)))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExt.Text = "", Null, Trim(txtKdExt.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, IIf(txtNamaExt.Text = "", Null, Trim(txtNamaExt.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_KategoryPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_KategoryPegawai = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_DetailKategoryPegawai(f_status As String) As Boolean
    sp_DetailKategoryPegawai = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDetailKategoryPegawai", adVarChar, adParamInput, 2, txtKdDetail.Text)
        .Parameters.Append .CreateParameter("DetailKategoryPegawai", adVarChar, adParamInput, 50, Trim(txtNamaDetail.Text))
        .Parameters.Append .CreateParameter("ReportDisplay", adVarChar, adParamInput, 50, IIf(txtRepDisplayDetail.Text = "", Null, Trim(txtRepDisplayDetail.Text)))
        .Parameters.Append .CreateParameter("KdKategoryPegawai", adChar, adParamInput, 1, IIf(dcKategory.Text = "", Null, Trim(dcKategory.BoundText)))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExtDetail.Text = "", Null, Trim(txtKdExtDetail.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, IIf(txtNamaExtDetail.Text = "", Null, Trim(txtNamaExtDetail.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl1.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_DetailKategoryPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_DetailKategoryPegawai = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub cmdCancel_Click()
    Call blankfield
    Call Dag
    Call SSTab1_KeyPress(13)
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    Select Case SSTab1.Tab
        Case 0
            If dgKategory.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakKategoriPegawai.Show
        Case 1
            If dgDetailKategory.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakDetailKategoriPegawai.Show
    End Select
hell:
End Sub

Private Sub cmdDel_Click()
    On Error GoTo hell

    If MsgBox("Yakin akan menghapus data ini?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case SSTab1.Tab
        Case 0
'            If Periksa("text", txtKategory, "Nama kategory pegawai kosong") = False Then Exit Sub
'            If sp_KategoryPegawai("D") = False Then Exit Sub
            
            If Periksa("text", txtKategory, "Nama kategory pegawai kosong") = False Then Exit Sub
            strSQL = "select * from DetailKategoryPegawai where KdKategoryPegawai='" & txtKdKategory & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                If sp_KategoryPegawai("D") = False Then Exit Sub
            End If
            
        Case 1
'            If Periksa("text", txtNamaDetail, "Detail kategory pegawai kosong") = False Then Exit Sub
'            If sp_DetailKategoryPegawai("D") = False Then Exit Sub
            
            If Periksa("text", txtNamaDetail, "Detail kategory pegawai kosong") = False Then Exit Sub
            strSQL = "select * from DataCurrentPegawai where KdDetailKategoryPegawai='" & txtKdDetail & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                If sp_DetailKategoryPegawai("D") = False Then Exit Sub
            End If
    End Select

    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call cmdCancel_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errLoad

    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtKategory, "Silahkan isi nama kategory pegawai ") = False Then Exit Sub
            If sp_KategoryPegawai("A") = False Then Exit Sub

        Case 1
            If dcKategory.Text <> "" Then
                If Periksa("datacombo", dcKategory, "Kategory Pegawai Tidak Terdaftar") = False Then Exit Sub
            End If
            
            If Periksa("text", txtNamaDetail, "Silahkan isi nama detail kategory pegawai ") = False Then Exit Sub
            If sp_DetailKategoryPegawai("A") = False Then Exit Sub
    End Select

    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call cmdCancel_Click
    txtKategory.SetFocus

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKategory_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtNamaDetail.SetFocus

On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKategory.Text)) = 0 Then txtNamaDetail.SetFocus: Exit Sub
        If dcKategory.MatchedWithList = True Then txtNamaDetail.SetFocus: Exit Sub
        strSQL = "SELECT KdKategoryPegawai, KategoryPegawai FROM KategoryPegawai WHERE (KategoryPegawai LIKE '%" & dcKategory.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKategory.BoundText = rs(0).Value
        dcKategory.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgDetailKategory_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDetailKategory
    WheelHook.WheelHook dgDetailKategory
End Sub

Private Sub dgDetailKategory_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaDetail.SetFocus
End Sub

Private Sub dgDetailKategory_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgDetailKategory.ApproxCount = 0 Then Exit Sub
    txtKdDetail = dgDetailKategory.Columns(0)
    txtNamaDetail = dgDetailKategory.Columns(1)
    If IsNull(dgDetailKategory.Columns(2)) Then txtRepDisplayDetail.Text = "" Else txtRepDisplayDetail = dgDetailKategory.Columns(2)
    If IsNull(dgDetailKategory.Columns(7)) Then dcKategory.BoundText = "" Else dcKategory.BoundText = dgDetailKategory.Columns(7)
    If IsNull(dgDetailKategory.Columns(3)) Then txtKdExtDetail.Text = "" Else txtKdExtDetail.Text = dgDetailKategory.Columns(3)
    If IsNull(dgDetailKategory.Columns(4)) Then txtNamaExtDetail.Text = "" Else txtNamaExtDetail.Text = dgDetailKategory.Columns(4)
    CheckStatusEnbl1.Value = dgDetailKategory.Columns(5).Value
End Sub

Private Sub dgKategory_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKategory
    WheelHook.WheelHook dgKategory
End Sub

Private Sub dgKategory_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKategory.SetFocus
End Sub

Private Sub dgKategory_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgKategory.ApproxCount = 0 Then Exit Sub
    txtKdKategory.Text = dgKategory.Columns(0).Value
    txtKategory.Text = dgKategory.Columns(1)
    If IsNull(dgKategory.Columns(2)) Then txtRepDisplay.Text = "" Else txtRepDisplay.Text = dgKategory.Columns(2)
    If IsNull(dgKategory.Columns(3)) Then txtKdExt.Text = "" Else txtKdExt.Text = dgKategory.Columns(3)
    If IsNull(dgKategory.Columns(4)) Then txtNamaExt.Text = "" Else txtNamaExt.Text = dgKategory.Columns(4)
    CheckStatusEnbl.Value = dgKategory.Columns(5).Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKey1
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 0
        Case vbKey2
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 1
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call cmdCancel_Click
    SSTab1.Tab = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call cmdCancel_Click
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case SSTab1.Tab
            Case 0
                txtKategory.SetFocus
            Case 1
                dcKategory.SetFocus
        End Select
    End If
errLoad:
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExt.SetFocus
End Sub

Private Sub txtKdExtDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExtDetail.SetFocus
End Sub

Private Sub txtKdKategory_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKategory.SetFocus
End Sub

Private Sub txtKdDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaDetail.SetFocus
End Sub

Private Sub txtNamaDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgDetailKategory.SetFocus
    End Select
End Sub

Private Sub txtNamaDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRepDisplayDetail.SetFocus
End Sub

Private Sub txtKategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgKategory.SetFocus
    End Select
End Sub

Private Sub txtKategory_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRepDisplay.SetFocus
End Sub

Private Sub txtNamaExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtNamaExtDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtParameter_Change()
    Call Dag
    strCetak = " where KategoryPegawai LIKE '%" & txtParameter.Text & "%'"
End Sub

Private Sub txtParameterDetail_Change()
    Call Dag
    strCetak = " where [Nama Detail] LIKE '%" & txtParameterDetail.Text & "%'"
End Sub

Private Sub txtRepDisplay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtRepDisplayDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtDetail.SetFocus
End Sub

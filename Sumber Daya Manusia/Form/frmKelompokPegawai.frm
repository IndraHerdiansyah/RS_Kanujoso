VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmKelompokPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kelompok & Detail Kelompok Pegawai"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKelompokPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   9150
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   2280
      TabIndex        =   31
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3600
      TabIndex        =   27
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4920
      TabIndex        =   28
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6240
      TabIndex        =   29
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   7560
      TabIndex        =   30
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
      TabIndex        =   13
      Top             =   1080
      Width           =   8940
      Begin TabDlg.SSTab SSTab1 
         Height          =   6570
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8700
         _ExtentX        =   15346
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
         TabCaption(0)   =   "Kelompok Pegawai"
         TabPicture(0)   =   "frmKelompokPegawai.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail Kelompok Pegawai"
         TabPicture(1)   =   "frmKelompokPegawai.frx":0CE6
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
            Height          =   6135
            Left            =   -74880
            TabIndex        =   17
            Top             =   360
            Width           =   8400
            Begin VB.TextBox txtParameter 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   4560
               MaxLength       =   50
               TabIndex        =   34
               Top             =   5680
               Width           =   3735
            End
            Begin VB.CheckBox chkStatus 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               Height          =   255
               Left            =   5880
               TabIndex        =   32
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
               MaxLength       =   50
               TabIndex        =   3
               Top             =   1080
               Width           =   5055
            End
            Begin VB.TextBox txtKdKelompok 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   2160
               MaxLength       =   2
               TabIndex        =   1
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txtKelompok 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   2
               Top             =   720
               Width           =   5055
            End
            Begin MSDataGridLib.DataGrid dgKelompok 
               Height          =   3315
               Left            =   135
               TabIndex        =   6
               Top             =   2280
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   5847
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
               Caption         =   "Cari Nama Kelompok"
               Height          =   210
               Left            =   2760
               TabIndex        =   35
               Top             =   5760
               Width           =   1650
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nama external"
               Height          =   210
               Left            =   480
               TabIndex        =   24
               Top             =   1860
               Width           =   1170
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   480
               TabIndex        =   23
               Top             =   1500
               Width           =   1140
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Report Display"
               Height          =   210
               Left            =   480
               TabIndex        =   22
               Top             =   1140
               Width           =   1155
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Nama Kelompok"
               Height          =   210
               Left            =   480
               TabIndex        =   19
               Top             =   765
               Width           =   1305
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Kode Kelompok"
               Height          =   210
               Left            =   480
               TabIndex        =   18
               Top             =   390
               Width           =   1275
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
            Height          =   6075
            Left            =   120
            TabIndex        =   14
            Top             =   405
            Width           =   8460
            Begin VB.TextBox txtParameterDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   4920
               MaxLength       =   50
               TabIndex        =   36
               Top             =   5640
               Width           =   3360
            End
            Begin VB.CheckBox chkStatus1 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               Height          =   255
               Left            =   6240
               TabIndex        =   33
               Top             =   1440
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtNamaExtDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2400
               MaxLength       =   30
               TabIndex        =   11
               Top             =   1800
               Width           =   5160
            End
            Begin VB.TextBox txtKdExtDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2400
               MaxLength       =   15
               TabIndex        =   10
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox txtNamaDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2400
               MaxLength       =   50
               TabIndex        =   9
               Top             =   1080
               Width           =   5160
            End
            Begin VB.TextBox txtKdDetail 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   2400
               MaxLength       =   2
               TabIndex        =   7
               Top             =   360
               Width           =   960
            End
            Begin MSDataGridLib.DataGrid dgDetailKelompok 
               Height          =   3255
               Left            =   120
               TabIndex        =   12
               Top             =   2280
               Width           =   8235
               _ExtentX        =   14526
               _ExtentY        =   5741
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
            Begin MSDataListLib.DataCombo dcKelompok 
               Height          =   330
               Left            =   2400
               TabIndex        =   8
               Top             =   720
               Width           =   3360
               _ExtentX        =   5927
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
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Cari Nama Detail Kelompok"
               Height          =   210
               Left            =   2520
               TabIndex        =   37
               Top             =   5700
               Width           =   2170
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Nama External"
               Height          =   210
               Left            =   480
               TabIndex        =   26
               Top             =   1860
               Width           =   1170
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   480
               TabIndex        =   25
               Top             =   1500
               Width           =   1140
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Kelompok Pegawai"
               Height          =   210
               Left            =   480
               TabIndex        =   20
               Top             =   765
               Width           =   1530
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Detail Kelompok"
               Height          =   210
               Left            =   480
               TabIndex        =   16
               Top             =   1140
               Width           =   1815
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Kode Detail"
               Height          =   210
               Left            =   480
               TabIndex        =   15
               Top             =   405
               Width           =   930
            End
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   21
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
      Left            =   7320
      Picture         =   "frmKelompokPegawai.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKelompokPegawai.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmKelompokPegawai.frx":30E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmKelompokPegawai.frx":5AA9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmKelompokPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub blankfield()
    On Error Resume Next
    Select Case SSTab1.Tab
        Case 0
            txtKdKelompok.Text = ""
            txtKelompok.Text = ""
            txtRepDisplay.Text = ""
            txtKdExt.Text = ""
            txtNamaExt.Text = ""
            txtKelompok.SetFocus
            chkStatus.Value = 1
        Case 1
            dcKelompok = ""
            txtKdDetail = ""
            txtNamaDetail = ""
            txtKdExtDetail.Text = ""
            txtNamaExtDetail.Text = ""
            dcKelompok.SetFocus
            chkStatus1.Value = 1
    End Select
End Sub

Sub Dag()
    On Error GoTo Errload
    Select Case SSTab1.Tab
        Case 0
            strSQL = "SELECT KdKelompokPegawai AS Kode, KelompokPegawai AS [Nama Kelompok], ReportDisplay AS [Report Display], KodeExternal AS [Kd.Ext], NamaExternal AS [Nama Ext], StatusEnabled AS Status FROM KelompokPegawai where KelompokPegawai LIKE '%" & txtParameter.Text & "%'" 'WHERE (StatusEnabled <> 0) OR (StatusEnabled IS NULL)"
            Call msubRecFO(rs, strSQL)
            Set dgKelompok.DataSource = rs
            dgKelompok.Columns(1).Width = 3000
            dgKelompok.Columns(5).Width = 1000

        Case 1
            strSQL = "SELECT * from V_KelompokPegawai where [Nama Detail] LIKE '%" & txtParameterDetail.Text & "%' "
            Call msubRecFO(rs, strSQL)
            Set dgDetailKelompok.DataSource = rs
            dgDetailKelompok.Columns(1).Width = 3000
            dgDetailKelompok.Columns(2).Width = 0
            dgDetailKelompok.Columns(3).Width = 3000
            dgDetailKelompok.Columns(6).Width = 1000

            Call msubDcSource(dcKelompok, rs, "SELECT KdKelompokPegawai, KelompokPegawai FROM KelompokPegawai where statusenabled='1'")
    End Select
    Exit Sub
Errload:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Function sp_KelompokPegawai(f_Status As String) As Boolean
On Error GoTo ErrspKelompokPegawai
    sp_KelompokPegawai = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKelompokPegawai", adChar, adParamInput, 2, txtKdKelompok.Text)
        .Parameters.Append .CreateParameter("KelompokPegawai", adVarChar, adParamInput, 50, Trim(txtKelompok.Text))
        .Parameters.Append .CreateParameter("ReportDisplay", adVarChar, adParamInput, 50, IIf(txtRepDisplay.Text = "", Null, Trim(txtRepDisplay.Text)))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExt.Text = "", Null, Trim(txtKdExt.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, IIf(txtNamaExt.Text = "", Null, Trim(txtNamaExt.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStatus.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_KelompokPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_KelompokPegawai = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
ErrspKelompokPegawai:
    If f_Status = "D" Then
            MsgBox "Hapus Data gagal, Data Sudah Dipakai", vbCritical
      Else
            Call msubPesanError
    End If
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    sp_KelompokPegawai = False
    
End Function

Private Function sp_DetailKelompokPegawai(f_Status As String) As Boolean
On Error GoTo ErrspDetailKelompokPegawai
    sp_DetailKelompokPegawai = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDetailKelompokPegawai", adChar, adParamInput, 2, txtKdDetail.Text)
        .Parameters.Append .CreateParameter("DetailKelompokPegawai", adVarChar, adParamInput, 50, Trim(txtNamaDetail.Text))
        .Parameters.Append .CreateParameter("KdKelompokPegawai", adChar, adParamInput, 2, IIf(dcKelompok.Text = "", Null, Trim(dcKelompok.BoundText)))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExtDetail.Text = "", Null, Trim(txtKdExtDetail.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, IIf(txtNamaExtDetail.Text = "", Null, Trim(txtNamaExtDetail.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStatus1.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_DetailKelompokPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_DetailKelompokPegawai = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
ErrspDetailKelompokPegawai:
    If f_Status = "D" Then
            MsgBox "Hapus Data gagal, Data Sudah Dipakai", vbCritical
      Else
            Call msubPesanError
    End If
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    sp_DetailKelompokPegawai = False
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
            If dgKelompok.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakKelompokPegawai.Show
        Case 1
            If dgDetailKelompok.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakDetailKelompokPegawai.Show
    End Select
hell:
End Sub

Private Sub cmdDel_Click()
    On Error GoTo hell

    If MsgBox("Yakin akan menghapus data ini?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtKelompok, "Nama kelompok pegawai kosong") = False Then Exit Sub
            If sp_KelompokPegawai("D") = False Then Exit Sub
        Case 1
            If Periksa("text", txtNamaDetail, "Detail kelompok pegawai kosong") = False Then Exit Sub
            If sp_DetailKelompokPegawai("D") = False Then Exit Sub
    End Select

    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    Call cmdCancel_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errload

    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtKelompok, "Silahkan isi nama kelompok ") = False Then Exit Sub
            If sp_KelompokPegawai("A") = False Then Exit Sub

        Case 1
            If dcKelompok.Text <> "" Then
                If Periksa("datacombo", dcKelompok, "Kelompok Pegawai Tidak Terdaftar") = False Then Exit Sub
            End If
            
            If Periksa("text", txtNamaDetail, "Silahkan isi nama detail kelompok ") = False Then Exit Sub
            If sp_DetailKelompokPegawai("A") = False Then Exit Sub

    End Select

    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    Call cmdCancel_Click

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcKelompok_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtNamaDetail.SetFocus

On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKelompok.Text)) = 0 Then txtNamaDetail.SetFocus: Exit Sub
        If dcKelompok.MatchedWithList = True Then txtNamaDetail.SetFocus: Exit Sub
        strSQL = "SELECT KdKelompokPegawai, KelompokPegawai FROM KelompokPegawai WHERE (KelompokPegawai LIKE '%" & dcKelompok.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKelompok.BoundText = rs(0).Value
        dcKelompok.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgDetailKelompok_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDetailKelompok
    WheelHook.WheelHook dgDetailKelompok
End Sub

Private Sub dgDetailKelompok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaDetail.SetFocus
End Sub

Private Sub dgDetailKelompok_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Errload
    If dgDetailKelompok.ApproxCount = 0 Then Exit Sub
    txtKdDetail = dgDetailKelompok.Columns(0)
    txtNamaDetail = dgDetailKelompok.Columns(1)
    If IsNull(dgDetailKelompok.Columns(2)) Then dcKelompok.BoundText = "" Else dcKelompok.BoundText = dgDetailKelompok.Columns(2)
    If IsNull(dgDetailKelompok.Columns(4)) Then txtKdExtDetail.Text = "" Else txtKdExtDetail.Text = dgDetailKelompok.Columns(4)
    If IsNull(dgDetailKelompok.Columns(5)) Then txtNamaDetail.Text = "" Else txtNamaExtDetail.Text = dgDetailKelompok.Columns(5)
    chkStatus1.Value = dgDetailKelompok.Columns("Status").Value
    Exit Sub
Errload:
End Sub

Private Sub dgKelompok_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKelompok
    WheelHook.WheelHook dgKelompok
End Sub

Private Sub dgKelompok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKelompok.SetFocus
End Sub

Private Sub dgKelompok_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgKelompok.ApproxCount = 0 Then Exit Sub
    txtKdKelompok.Text = dgKelompok.Columns(0)
    txtKelompok.Text = dgKelompok.Columns(1)
    If IsNull(dgKelompok.Columns(2)) Then txtRepDisplay.Text = "" Else txtRepDisplay.Text = dgKelompok.Columns(2)
    If IsNull(dgKelompok.Columns(3)) Then txtKdExt.Text = "" Else txtKdExt.Text = dgKelompok.Columns(3)
    If IsNull(dgKelompok.Columns(4)) Then txtNamaExt.Text = "" Else txtNamaExt.Text = dgKelompok.Columns(4)
    chkStatus.Value = dgKelompok.Columns(5).Value
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
    On Error GoTo Errload
    If KeyAscii = 13 Then
        Select Case SSTab1.Tab
            Case 0
            
            Case 1
        
        End Select
    End If
Errload:
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExt.SetFocus
End Sub

Private Sub txtKdExtDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExtDetail.SetFocus
End Sub

Private Sub txtKdKelompok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKelompok.SetFocus
End Sub

Private Sub txtKdDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaDetail.SetFocus
End Sub

Private Sub txtNamaDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgDetailKelompok.SetFocus
    End Select
End Sub

Private Sub txtNamaDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtDetail.SetFocus
End Sub

Private Sub txtKelompok_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgKelompok.SetFocus
    End Select
End Sub

Private Sub txtKelompok_KeyPress(KeyAscii As Integer)
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
    strCetak = " where KelompokPegawai LIKE '%" & txtParameter.Text & "%'"
End Sub

Private Sub txtParameterDetail_Change()
    Call Dag
    strCetak = " where [Nama Detail] LIKE '%" & txtParameterDetail.Text & "%'"
End Sub

Private Sub txtRepDisplay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtRepDisplayDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKelompok.SetFocus
End Sub

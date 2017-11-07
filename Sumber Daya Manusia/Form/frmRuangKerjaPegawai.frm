VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRuangKerjaPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Ruang Kerja & Sub Ruang Kerja Pegawai"
   ClientHeight    =   8790
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
   Icon            =   "frmRuangKerjaPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Baru"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   7440
      TabIndex        =   20
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
      TabIndex        =   8
      Top             =   1080
      Width           =   8700
      Begin TabDlg.SSTab SSTab1 
         Height          =   6570
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   11589
         _Version        =   393216
         Tabs            =   2
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
         TabCaption(0)   =   "Ruang Kerja"
         TabPicture(0)   =   "frmRuangKerjaPegawai.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Sub Ruang Kerja"
         TabPicture(1)   =   "frmRuangKerjaPegawai.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
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
            Height          =   5775
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   7920
            Begin VB.CheckBox chkStsRuangan 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4440
               TabIndex        =   30
               Top             =   1080
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.TextBox txtNmExt 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1800
               MaxLength       =   50
               TabIndex        =   22
               Top             =   1440
               Width           =   4215
            End
            Begin VB.TextBox txtKdExt 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1800
               MaxLength       =   15
               TabIndex        =   21
               Top             =   1080
               Width           =   2295
            End
            Begin VB.TextBox txtKdRuangKerja 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   1800
               MaxLength       =   3
               TabIndex        =   1
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtRuangKerja 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1800
               MaxLength       =   30
               TabIndex        =   2
               Top             =   720
               Width           =   5055
            End
            Begin MSDataGridLib.DataGrid dgRuangKerja 
               Height          =   3690
               Left            =   135
               TabIndex        =   3
               Top             =   1920
               Width           =   7560
               _ExtentX        =   13335
               _ExtentY        =   6509
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
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Nama External"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   24
               Top             =   1440
               Width           =   1155
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   23
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Ruang Kerja"
               Height          =   210
               Left            =   120
               TabIndex        =   14
               Top             =   735
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Kode Ruang"
               Height          =   210
               Left            =   120
               TabIndex        =   13
               Top             =   360
               Width           =   990
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
            Height          =   5715
            Left            =   -74760
            TabIndex        =   9
            Top             =   480
            Width           =   7980
            Begin VB.CheckBox chkSts 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4440
               TabIndex        =   29
               Top             =   1320
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.TextBox txtNmExtKerja 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1800
               MaxLength       =   50
               TabIndex        =   26
               Top             =   1680
               Width           =   4215
            End
            Begin VB.TextBox txtKdExtKerja 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1800
               MaxLength       =   15
               TabIndex        =   25
               Top             =   1320
               Width           =   2295
            End
            Begin VB.TextBox txtNamaDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1800
               MaxLength       =   50
               TabIndex        =   6
               Top             =   960
               Width           =   5040
            End
            Begin VB.TextBox txtKdDetail 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   1800
               MaxLength       =   3
               TabIndex        =   4
               Top             =   240
               Width           =   840
            End
            Begin MSDataGridLib.DataGrid dgSubRuangKerja 
               Height          =   3360
               Left            =   120
               TabIndex        =   7
               Top             =   2160
               Width           =   7680
               _ExtentX        =   13547
               _ExtentY        =   5927
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
            Begin MSDataListLib.DataCombo dcRuangKerja 
               Height          =   330
               Left            =   1800
               TabIndex        =   5
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
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Nama External"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   28
               Top             =   1680
               Width           =   1155
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   27
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Ruang Kerja"
               Height          =   210
               Left            =   120
               TabIndex        =   15
               Top             =   645
               Width           =   975
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Sub Ruangan"
               Height          =   210
               Left            =   120
               TabIndex        =   11
               Top             =   1020
               Width           =   1590
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Kode Sub"
               Height          =   210
               Left            =   120
               TabIndex        =   10
               Top             =   285
               Width           =   795
            End
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   16
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
      Picture         =   "frmRuangKerjaPegawai.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRuangKerjaPegawai.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRuangKerjaPegawai.frx":30E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmRuangKerjaPegawai.frx":5AA9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmRuangKerjaPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lcStatus As String

Sub blankfield()
On Error Resume Next
    Select Case SSTab1.Tab
    Case 0
        txtKdRuangKerja.Text = ""
        txtRuangKerja.Text = ""
        txtKdExt = ""
        txtNmExt = ""
        txtRuangKerja.SetFocus
    Case 1
        dcRuangKerja = ""
        txtKdDetail = ""
        txtNamaDetail = ""
        txtKdExtKerja = ""
        txtNmExtKerja = ""
        txtNamaDetail.SetFocus
    End Select
End Sub

Sub Dag()
On Error GoTo errLoad
    Select Case SSTab1.Tab
        Case 0
            strsql = "SELECT * FROM RuangKerja"
            Call msubRecFO(rs, strsql)
            Set dgRuangKerja.DataSource = rs
                dgRuangKerja.Columns(0).Width = 1000
                dgRuangKerja.Columns(1).Width = 3000
                dgRuangKerja.Columns(2).Width = 1000
                dgRuangKerja.Columns(3).Width = 3000
                dgRuangKerja.Columns(4).Width = 1000
        
        Case 1
            strsql = "SELECT * from SubRuangKerja"
            Call msubRecFO(rs, strsql)
            Set dgSubRuangKerja.DataSource = rs
                dgSubRuangKerja.Columns(0).Width = 1000
                dgSubRuangKerja.Columns(1).Width = 3000
                dgSubRuangKerja.Columns(2).Width = 0
                dgSubRuangKerja.Columns(3).Width = 1000
                dgSubRuangKerja.Columns(4).Width = 3000
                dgSubRuangKerja.Columns(5).Width = 1000
            
            Call msubDcSource(dcRuangKerja, rs, "SELECT * FROM RuangKerja WHERE StatusEnabled=1")
    End Select
Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Function sp_RuangKerja(f_status As String) As Boolean
sp_RuangKerja = True
Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdRuangKerja", adChar, adParamInput, 3, txtKdRuangKerja.Text)
        .Parameters.Append .CreateParameter("RuangKerja", adVarChar, adParamInput, 50, Trim(txtRuangKerja.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExt.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExt.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsRuangan.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_RuangKerja"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_RuangKerja = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_SubRuangKerja(f_status As String) As Boolean
sp_SubRuangKerja = True
Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdSubRuangKerja", adVarChar, adParamInput, 3, txtKdDetail.Text)
        .Parameters.Append .CreateParameter("SubRuangKerja", adVarChar, adParamInput, 50, Trim(txtNamaDetail.Text))
        .Parameters.Append .CreateParameter("KdRuangKerja", adChar, adParamInput, 3, IIf(dcRuangKerja.Text = "", Null, Trim(dcRuangKerja.BoundText)))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtKerja.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtKerja.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_SubRuangKerja"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_SubRuangKerja = False
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

Private Sub cmdDel_Click()
On Error GoTo hell
    
    If MsgBox("Yakin akan menghapus data ini?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtRuangKerja, "Nama ruang kerja pegawai kosong") = False Then Exit Sub
            If sp_RuangKerja("D") = False Then Exit Sub
        Case 1
            If Periksa("text", txtNamaDetail, "Sub ruang kerja pegawai kosong") = False Then Exit Sub
            If sp_SubRuangKerja("D") = False Then Exit Sub
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
            If Periksa("text", txtRuangKerja, "Silahkan isi nama ruang kerja pegawai ") = False Then Exit Sub
            If sp_RuangKerja("A") = False Then Exit Sub

        Case 1
            If Periksa("text", txtNamaDetail, "Silahkan isi nama sub ruang kerja pegawai ") = False Then Exit Sub
            If sp_SubRuangKerja("A") = False Then Exit Sub
    End Select
    
    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call cmdCancel_Click
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcRuangKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaDetail.SetFocus
End Sub

Private Sub dgSubRuangKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaDetail.SetFocus
End Sub

Private Sub dgSubRuangKerja_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgSubRuangKerja.ApproxCount = 0 Then Exit Sub
    txtKdDetail = dgSubRuangKerja.Columns(0)
    txtNamaDetail = dgSubRuangKerja.Columns(1)
    If IsNull(dgSubRuangKerja.Columns(2)) Then dcRuangKerja.BoundText = "" Else dcRuangKerja.BoundText = dgSubRuangKerja.Columns(2)
    txtKdExtKerja.Text = dgSubRuangKerja.Columns(3)
    txtNmExtKerja.Text = dgSubRuangKerja.Columns(4)
    
    If dgSubRuangKerja.Columns(5).Value = "<Type mismacth>" Then
        chkSts.Value = 0
    Else
        If dgSubRuangKerja.Columns(5).Value = 1 Then
            chkSts.Value = 1
        Else
            chkSts.Value = 0
        End If
    End If
    
End Sub

Private Sub dgRuangKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRuangKerja.SetFocus
End Sub

Private Sub dgRuangKerja_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgRuangKerja.ApproxCount = 0 Then Exit Sub
    txtKdRuangKerja.Text = dgRuangKerja.Columns(0).Value
    txtRuangKerja.Text = dgRuangKerja.Columns(1)
    txtKdExt.Text = dgRuangKerja.Columns(2)
    txtNmExt.Text = dgRuangKerja.Columns(3)
    
    If dgRuangKerja.Columns(4).Value = "<Type mismacth>" Then
        chkStsRuangan.Value = 0
    Else
        If dgRuangKerja.Columns(4).Value = 1 Then
            chkStsRuangan.Value = 1
        Else
            chkStsRuangan.Value = 0
        End If
    End If
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
    Call Dag
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case SSTab1.Tab
            Case 0
                txtRuangKerja.SetFocus
            Case 1
                dcRuangKerja.SetFocus
        End Select
    End If
errLoad:
End Sub
    
Private Sub txtKdRuangKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRuangKerja.SetFocus
End Sub

Private Sub txtKdDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaDetail.SetFocus
End Sub

Private Sub txtNamaDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgSubRuangKerja.SetFocus
    End Select
End Sub

Private Sub txtNamaDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtRuangKerja_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgRuangKerja.SetFocus
    End Select
End Sub

Private Sub txtRuangKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

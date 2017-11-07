VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJadwalKerja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Jadwal Kerja Pegawai"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   Icon            =   "frmJadwalKerja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11115
   Begin VB.CommandButton cmdRefresh 
      Height          =   375
      Left            =   3120
      Picture         =   "frmJadwalKerja.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Refresh Jadwal"
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid fgJadwalKerja 
      Height          =   2895
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Visible         =   0   'False
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5106
      _Version        =   393216
      FocusRect       =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkTabelJadwal 
      Caption         =   "&Munculkan Tabel"
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
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8040
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8505
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   19553
            Text            =   "F1 - Cetak"
            TextSave        =   "F1 - Cetak"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOtomatis 
      Caption         =   "&Jadwal Automatis"
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
      TabIndex        =   21
      Top             =   8040
      Width           =   1455
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
      Left            =   8040
      TabIndex        =   17
      Top             =   8040
      Width           =   1455
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
      Left            =   9600
      TabIndex        =   16
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox txtnamapegawai 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7920
      TabIndex        =   14
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kalendar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   3975
      Begin VB.TextBox lblKeterangan 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   3060
         Width           =   3735
      End
      Begin MSFlexGridLib.MSFlexGrid fgKalender 
         Height          =   2415
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4260
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   16384
         ForeColorFixed  =   16777215
         WordWrap        =   -1  'True
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   109510659
         UpDown          =   -1  'True
         CurrentDate     =   39554
      End
      Begin MSDataGridLib.DataGrid dgTanggal 
         Height          =   2415
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   4260
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
      Begin VB.Label lblKeterangan1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   3735
      End
   End
   Begin VB.TextBox txtidpegawai 
      Height          =   285
      Left            =   7680
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtharilibur 
      Height          =   285
      Left            =   5760
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtnamatgl 
      Height          =   285
      Left            =   4920
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtkdtgl 
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid dgJadwalKerja 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSDataListLib.DataCombo dcTempatBertugas 
      Height          =   330
      Left            =   4200
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
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
   Begin MSDataListLib.DataCombo dcShiftKerja 
      Height          =   315
      Left            =   9240
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin MSComctlLib.ListView lvjadwalkerja 
      Height          =   2775
      Left            =   4200
      TabIndex        =   7
      Top             =   1680
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4895
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Diagnosa"
         Object.Width           =   13229
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   25
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
   Begin MSDataListLib.DataCombo dcShiftKerja2 
      Height          =   315
      Left            =   7320
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Shift Kerja"
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
      Left            =   7320
      TabIndex        =   26
      Top             =   1080
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cari Nama Pegawai :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6120
      TabIndex        =   15
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Jam Kerja"
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
      Left            =   9240
      TabIndex        =   6
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ruangan Tempat Bertugas"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   1080
      Width           =   1920
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9240
      Picture         =   "frmJadwalKerja.frx":0E14
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmJadwalKerja.frx":1B9C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmJadwalKerja.frx":455D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmJadwalKerja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intRowLama As Integer, intColLama As Integer
Private intRowCur As Integer, intColCur As Integer
Private blnHariLibur As Boolean
Private intTglLibur As Integer

Dim i, j As Integer
Dim strKey As String

Private Sub chkTabelJadwal_Click()
    If Me.chkTabelJadwal.Value = 1 Then
        Call subSetFGJadwalKerja
        Call subIsiFgJadwalKerja
        Me.dcShiftKerja.Enabled = False
        Me.fgJadwalKerja.Visible = True
        Me.cmdRefresh.Visible = True
        
        Me.dcShiftKerja2.Enabled = False
        Me.dcTempatBertugas.Enabled = False
        Me.lvjadwalkerja.Enabled = False
        Me.txtnamapegawai.Enabled = False
    Else
        Call subdgJadwalKerja
        Me.fgJadwalKerja.Visible = False
        Me.fgJadwalKerja.clear
        Me.dcShiftKerja.Enabled = True
        Me.cmdRefresh.Visible = False
        
        Me.dcShiftKerja2.Enabled = True
        Me.dcTempatBertugas.Enabled = True
        Me.lvjadwalkerja.Enabled = True
        Me.txtnamapegawai.Enabled = True
    End If
End Sub

Private Sub cmdOtomatis_Click()
    glDate = Me.DTPicker1.Value
    With frmAutoJadwal
        .dcTempatBertugas.Text = Me.dcTempatBertugas.Text
        If .lvjadwalkerja.ListItems.Count > 0 Then
            .lvjadwalkerja.SetFocus
        End If
        .Show
    End With
End Sub

Private Sub cmdRefresh_Click()
    Call subSetFGJadwalKerja
    Call subIsiFgJadwalKerja
    Me.fgJadwalKerja.SetFocus
End Sub


Private Sub dcShiftKerja2_Change()
'    If dcShiftKerja2.Text = "Shift" Then
'        strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja_New order by NamaShift"
'        Call msubDcSource(dcShiftKerja, rs, strSQL)
'    Else
'        strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja_New where KdShift = '01'"
'        Call msubDcSource(dcShiftKerja, rs, strSQL)
'    End If
On Error GoTo hell

    Call loadListViewSource

    Call subLoadLView

    If dcShiftKerja2.MatchedWithList = False Then Exit Sub

    strSQL = "SELECT IdPegawai From ConvertIdPegawaiToShift WHERE (IdShift = '" & dcShiftKerja2.BoundText & "')"
    Call msubRecFO(rs, strSQL)
    For j = 0 To rs.RecordCount - 1
        For i = 1 To lvjadwalkerja.ListItems.Count
            If Right(lvjadwalkerja.ListItems(i).key, 10) = rs(0).Value Then
                lvjadwalkerja.ListItems.Item(i).Bold = True
                lvjadwalkerja.ListItems.Item(i).ForeColor = vbBlue
                lvjadwalkerja.ListItems.Item(i).Checked = True
            End If
        Next i
        rs.MoveNext
    Next j
    
    dcShiftKerja.Text = ""
    
    'strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja_New where NamaShift = '" & dcShiftKerja2.Text & "' order by NamaShift"
    strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja_New" ' where NamaShift = '" & dcShiftKerja2.Text & "' order by NamaShift"
    Call msubDcSource(dcShiftKerja, rs, strSQL)
    
    Exit Sub
hell:
    Call msubPesanError
    Set rs = Nothing
End Sub
'-------------------------------------
Private Sub subLoadLView()
    On Error GoTo hell

    strSQL = "SELECT DISTINCT IdPegawai, NamaLengkap From v_TempatBertugas WHERE NamaRuangan = '" & dcTempatBertugas & "' ORDER BY IdPegawai"

    Call msubRecFO(rs, strSQL)
    'lvjadwalkerja.ListItems.clear
    Do While rs.EOF = False
        strKey = "key" & rs(0).Value
        lvjadwalkerja.ListItems.add , strKey, rs(1).Value
        rs.MoveNext
    Loop
    lvjadwalkerja.View = lvwList
    Exit Sub
hell:
    Call msubPesanError
    Set rs = Nothing
End Sub
'-----------------------------------

Private Sub dcShiftKerja2_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcShiftKerja2.Text)) = 0 Then dcShiftKerja.SetFocus: Exit Sub
        If dcShiftKerja2.MatchedWithList = True Then dcShiftKerja.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "Select IdShift, Dinaskerja From DinasKerja WHERE DinasKerja LIKE '%" & dcShiftKerja2.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcShiftKerja2.BoundText = dbRst(0).Value
        dcShiftKerja2.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcTempatBertugas_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcTempatBertugas.Text)) = 0 Then dcShiftKerja2.SetFocus: Exit Sub
        If dcTempatBertugas.MatchedWithList = True Then dcShiftKerja2.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE NamaRuangan LIKE '%" & dcTempatBertugas.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcTempatBertugas.BoundText = dbRst(0).Value
        dcTempatBertugas.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub DTPicker1_Change()

    intRowCur = 0
    intColCur = 0
    intRowLama = 0
    intColLama = 0
    intTglLibur = 0
    blnHariLibur = False
    Me.DTPicker1.Day = 1
    Call subLoadKalender
    If Me.fgKalender.Rows = 1 Then
        Me.chkTabelJadwal.Enabled = False
        Exit Sub
    Else
        Me.chkTabelJadwal.Enabled = True
    End If
    If Me.fgJadwalKerja.Visible Then
        Call subSetFGJadwalKerja
        Call subIsiFgJadwalKerja
    End If
End Sub

Private Sub fgJadwalKerja_DblClick()
    Call fgJadwalKerja_KeyPress(13)
End Sub

Private Sub fgJadwalKerja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With Me.fgJadwalKerja
            .TextMatrix(.row, .Col) = ""
        End With
    End If
End Sub

Private Sub fgJadwalKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Me.fgJadwalKerja
            If .TextMatrix(.row, 0) = "" Then Exit Sub
            If .TextMatrix(1, .Col) = "" Then Exit Sub
            Select Case .TextMatrix(.row, .Col)
                Case ""
                    .TextMatrix(.row, .Col) = "P"
                Case "P"
                    .TextMatrix(.row, .Col) = "S"
                Case "S"
                    .TextMatrix(.row, .Col) = "M"
                Case "M"
                    .TextMatrix(.row, .Col) = "R"
                Case "R"
                    .TextMatrix(.row, .Col) = "L"
                Case "L"
                    .TextMatrix(.row, .Col) = ""
            End Select
'            Select Case .TextMatrix(.row, .Col)
'                Case ""
'                    .TextMatrix(.row, .Col) = "P"
'                Case "P"
'                    .TextMatrix(.row, .Col) = "S"
'                Case "S"
'                    .TextMatrix(.row, .Col) = "M"
'                Case "M"
'                    .TextMatrix(.row, .Col) = "L"
'                Case "L"
'                    .TextMatrix(.row, .Col) = ""
'            End Select
        End With
    End If
End Sub

Private Sub fgJadwalKerja_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Button = 1
        PopupMenu MDIUtama.mnuPopUp
    End If
End Sub

Private Sub fgKalender_Click()
    On Error GoTo tangani
    Dim strHari As String
    Dim strBuffer As String
    Dim sTgl, sBln, sThn As String

    With Me.fgKalender

        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        Me.lblKeterangan.Text = ""
        Me.DTPicker1.Day = .TextMatrix(.row, .Col)
        strSQL = "SELECT [Kode Tgl], Tanggal, [Hari Libur] FROM v_tanggal" & _
        " WHERE DAY(Tanggal)='" & Day(Me.DTPicker1.Value) & "'" & _
        " AND MONTH(Tanggal)='" & Month(Me.DTPicker1.Value) & "'" & _
        " AND YEAR(Tanggal)='" & Year(Me.DTPicker1.Value) & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount > 0 Then
'            Me.txtkdtgl.Text = rs.Fields.Item("Kode Tgl").Value
'            Me.txtnamatgl.Text = rs.Fields.Item("Tanggal").Value
'            Me.txtharilibur.Text = IIf(IsNull(rs.Fields.Item("Hari Libur").Value), "", rs.Fields.Item("Hari Libur").Value)
'            Me.lblKeterangan.Caption = IIf(Me.txtharilibur.Text <> "", FormatDateTime(Me.DTPicker1.Value, vbLongDate) & ": " & Me.txtharilibur.Text, "")
            
            Me.txtkdtgl.Text = rs.Fields.Item("Kode Tgl").Value
            Me.txtnamatgl.Text = rs.Fields.Item("Tanggal").Value
            strHari = WeekdayName(Weekday(DTPicker1.Value - 1))
            If strHari = "Sabtu" Or strHari = "Minggu" Then
                Me.txtharilibur.Text = "Hari " & strHari
                Me.lblKeterangan.Text = IIf(Me.txtharilibur.Text <> "", FormatDateTime(Me.DTPicker1.Value, vbLongDate) & ": " & Me.txtharilibur.Text, "")
            Else
            '//yayang.agus 2014-08-14
'                Me.txtharilibur.Text = IIf(IsNull(rs.Fields.Item("Hari Libur").Value), "", rs.Fields.Item("Hari Libur").Value)
                'Me.lblKeterangan.Text = IIf(Me.txtharilibur.Text <> "", FormatDateTime(Me.DTPicker1.Value, vbLongDate) & ": " & Me.txtharilibur.Text, "")
                Me.lblKeterangan.Text = FormatDateTime(Me.DTPicker1.Value, vbLongDate) & ": "
            End If
            If rs.RecordCount <> 0 Then
                For i = 0 To rs.RecordCount - 1
                    lblKeterangan.Text = lblKeterangan.Text & ", " & rs.Fields.Item("Hari Libur").Value
                    rs.MoveNext
                Next
            End If
            '//
        Else
            strHari = WeekdayName(Weekday(DTPicker1.Value - 1)) ', , vbUseSystemDayOfWeek)
            If strHari = "Sabtu" Or strHari = "Minggu" Then
                Me.txtharilibur.Text = "Hari " & strHari
                Me.lblKeterangan.Text = IIf(Me.txtharilibur.Text <> "", FormatDateTime(Me.DTPicker1.Value, vbLongDate) & ": " & Me.txtharilibur.Text, "")
            Else
                Me.txtkdtgl.Text = ""
                Me.txtnamatgl.Text = ""
                Me.txtharilibur.Text = ""
                Me.lblKeterangan.Text = ""
            End If
        End If

    End With
    If Me.fgJadwalKerja.Visible = True Then
        Call subTandaiTanggal(Me.DTPicker1.Value)
        Exit Sub
    End If
    Call loadListViewSource
    Exit Sub

tangani:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With frmLaporanJadwalKerja
            .Show
            .dcRuangan.Text = Me.dcTempatBertugas.Text
            .dtpAwal.Year = Me.DTPicker1.Year
            .dtpAwal.Month = Me.DTPicker1.Month
            .dtpAwal.Day = 1
            .dtpAkhir.Year = Me.DTPicker1.Year
            .dtpAkhir.Month = Me.DTPicker1.Month
            .dtpAkhir.Day = Me.DTPicker1.Day
        End With
    End If
End Sub

Public Sub loadListViewSource()
    On Error GoTo tangani

    strSQL = "SELECT DataCurrentPegawai.IdPegawai, DataPegawai.NamaLengkap " & _
    "FROM DataCurrentPegawai INNER JOIN " & _
    "DataPegawai ON DataCurrentPegawai.IdPegawai = DataPegawai.IdPegawai " & _
    "WHERE KdRuanganKerja = '" & dcTempatBertugas.BoundText & "' and KdStatus = '01' ORDER BY IdPegawai"
    
    Call msubRecFO(rs, strSQL)
    lvjadwalkerja.ListItems.clear
    While Not rs.EOF
        lvjadwalkerja.ListItems.add , "A" & rs(0).Value, rs(1).Value
        rs.MoveNext
    Wend
    lvjadwalkerja.Sorted = True

    If rs.RecordCount = 0 Then Exit Sub

    strSQL = "SELECT ID from V_JadwalKerjaNew WHERE KdShift = '" & dcShiftKerja.BoundText & "' AND KdRuangan = '" & dcTempatBertugas.BoundText & "' "

    Call msubRecFO(rs, strSQL)
    Do While rs.EOF = False
        lvjadwalkerja.ListItems("A" & rs(0)).Checked = True
        lvjadwalkerja.ListItems("A" & rs(0)).ForeColor = vbBlue
        lvjadwalkerja.ListItems("A" & rs(0)).Bold = True
        rs.MoveNext
    Loop
    Exit Sub

tangani:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Sub subDCSource()
    strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan order by NamaRuangan"
    Call msubDcSource(dcTempatBertugas, rs, strSQL)

    strSQL = "Select IdShift, Dinaskerja From DinasKerja"
    Call msubDcSource(dcShiftKerja2, rs, strSQL)

    strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja_New order by NamaShift"
    Call msubDcSource(dcShiftKerja, rs, strSQL)
End Sub

Private Sub cmdCetak_Click()
    frmCetakJadwalKerja.Show
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    If Me.fgJadwalKerja.Visible = True Then
        Call subSimpanJadwalBaru
        Exit Sub
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Function sp_JadwalKerja(f_KdPegawai As String) As Boolean
    sp_JadwalKerja = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglJadwalKerja", adDate, adParamInput, , DTPicker1)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_KdPegawai)
        .Parameters.Append .CreateParameter("KdShift", adChar, adParamInput, 2, dcShiftKerja.BoundText)
        
        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AU_JADWALKERJA"
        .Execute

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub subLoadGridTanggal()
    On Error GoTo errLoad
    strSQL = "SELECT [Kode Tgl], Tanggal, [Hari Libur] FROM v_tanggal" & _
    " WHERE MONTH(Tanggal)='" & Month(Me.DTPicker1.Value) & "'" & _
    " AND YEAR(Tanggal)='" & Year(Me.DTPicker1.Value) & "'"
    Call msubRecFO(rs, strSQL)
    Set dgTanggal.DataSource = rs
    With dgTanggal
        .Columns("Kode Tgl").Width = 0
        .Columns("Tanggal").Width = 1000
        .Columns("Hari Libur").Width = 2000
        .Visible = True
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcShiftKerja_Change()
    Call loadListViewSource
    Call subdgJadwalKerja
End Sub

Private Sub dcTempatBertugas_Change()
    Call loadListViewSource
    Call subdgJadwalKerja
    If Me.lvjadwalkerja.ListItems.Count > 0 Then
        If Me.fgKalender.Rows = 1 Then
            Me.chkTabelJadwal.Enabled = False
            Exit Sub
        Else
            Me.chkTabelJadwal.Enabled = True
        End If
        If Me.chkTabelJadwal.Value = 1 Then
            Call subSetFGJadwalKerja
            Call subIsiFgJadwalKerja
        End If
    Else
        Me.chkTabelJadwal.Enabled = False
        If Me.chkTabelJadwal.Value = 1 Then Me.chkTabelJadwal.Value = 0
    End If
End Sub

Private Sub dcTempatBertugas_GotFocus()
    On Error GoTo errLoad
    strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan Order By NamaRuangan"
    Call msubDcSource(dcTempatBertugas, rs, strSQL)
    Exit Sub
errLoad:
End Sub

Private Sub dgtanggal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgTanggal
        txtkdtgl.Text = .Columns("Kode Tgl").Value
        txtnamatgl.Text = .Columns("Tanggal").Value
        If .Columns("Hari Libur").Value = "" Then
            txtharilibur.Text = ""
        Else
            txtharilibur.Text = .Columns("Hari Libur").Value
        End If
        Call loadListViewSource
    End With
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    Me.DTPicker1.Value = Now
    Call subDCSource
    Call subLoadGridTanggal
    Call loadListViewSource
    Call subdgJadwalKerja
    Call subLoadKalender
End Sub

Private Sub lvjadwalkerja_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim intRowNama As Integer
    Dim intMaxCol As Integer, intMaxRow As Integer
    Dim i As Integer, j As Integer

    With Me.fgJadwalKerja
        intMaxCol = .Cols - 1
        intMaxRow = .Rows - 1

        For i = 2 To intMaxRow
            .Col = 0
            .row = i
            If .CellBackColor = &H808080 Then
                For j = 0 To intMaxCol
                    .Col = j
                    If .CellBackColor = vbYellow Then
                        .CellBackColor = &H808080
                    Else
                        .CellBackColor = vbWhite
                    End If
                Next
                .Col = 0
                .CellBackColor = &H8000000F
            End If
        Next
        intRowNama = funcCariRowNama(Item.Text)
        .row = intRowNama
        For i = 0 To intMaxCol
            .Col = i
            If .CellBackColor = &H808080 Then
                .CellBackColor = vbYellow
            Else
                .CellBackColor = &H808080
            End If
        Next
    End With
End Sub

Private Sub txtNamaPegawai_Change()
    strSQL = "SELECT DataPegawai.NamaLengkap, DataCurrentPegawai.IdPegawai  " & _
    "FROM DataCurrentPegawai INNER JOIN " & _
    "DataPegawai ON DataCurrentPegawai.IdPegawai = DataPegawai.IdPegawai " & _
    "WHERE KdRuanganKerja = '" & dcTempatBertugas.BoundText & "' AND KdStatus = '01' AND NamaLengkap LIKE '%" & txtnamapegawai & "%' ORDER BY IdPegawai"
    Call msubRecFO(rs, strSQL)
    lvjadwalkerja.ListItems.clear
    While Not rs.EOF
        lvjadwalkerja.ListItems.add , "A" & rs(1).Value, rs(0).Value
        rs.MoveNext
    Wend
    lvjadwalkerja.Sorted = True
    '----------------------------------
    strsqlx = "SELECT * FROM V_JadwalKerjaNew WHERE Ruangan='" & dcTempatBertugas & "' AND Nama LIKE '%" & txtnamapegawai & "%' ORDER BY Tanggal"
    Call msubRecFO(rsx, strsqlx)
    Set dgJadwalKerja.DataSource = rsx
    
    Exit Sub
End Sub

Private Sub subdgJadwalKerja()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "SELECT * FROM V_JadwalKerjaNew WHERE Ruangan='" & dcTempatBertugas & "'" & _
    " ORDER BY Tanggal"

    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgJadwalKerja.DataSource = rs
    With dgJadwalKerja
        .Columns("KdShift").Width = 0
        .Columns("Nama").Width = 2500
        .Columns("KdRuangan").Width = 0
    End With
    Exit Sub
    Set rs = Nothing
errLoad:
    Call msubPesanError
End Sub

'16 April 2008
Private Sub subSetFgKalender()
    Dim i As Integer
    Dim intLebarCell As Integer

    With Me.fgKalender
        .clear
        .Rows = 1
        .Cols = 7

        .RowHeight(0) = 500

        .TextMatrix(0, 0) = "Sen"
        .TextMatrix(0, 1) = "Sel"
        .TextMatrix(0, 2) = "Rab"
        .TextMatrix(0, 3) = "Kam"
        .TextMatrix(0, 4) = "Jum"
        .TextMatrix(0, 5) = "Sab"
        .TextMatrix(0, 6) = "Ming"
        .row = 0
        .Col = 5
        .CellBackColor = vbRed
        .Col = 6
        .CellBackColor = vbRed
        intLebarCell = (Me.fgKalender.Width / 7) - 2
        For i = 0 To 6
            .ColWidth(i) = intLebarCell
            .row = 0
            .Col = i
            .CellFontBold = True
        Next
    End With
End Sub

Private Sub subLoadKalender()

    Call subSetFgKalender
    Call DateSystem

End Sub

Private Sub DateSystem()

    Dim nmonth As Integer
    Dim nLastDay As Integer
    Dim nmodRemainder As Integer
    Dim calDate As Date

    calDate = Format(DTPicker1, "dd/mm/yyyy")
    nmonth = Month(calDate)

    If nmonth = 4 Or nmonth = 6 Or nmonth = 9 Or nmonth = 11 Then
        nLastDay = 30
    ElseIf nmonth = 2 Then
        nmodRemainder = Year(calDate) Mod 4
        If nmodRemainder = 0 Then
            nmodRemainder = Year(calDate) Mod 100
            If nmodRemainder = 0 Then
                nmodRemainder = Year(calDate) Mod 400
                If nmodRemainder = 0 Then
                    nLastDay = 29                ' Leap year
                Else
                    nLastDay = 28
                End If
            Else
                nLastDay = 29
            End If
        Else
            nLastDay = 28
        End If
    Else
        nLastDay = 31
    End If

    Dim strTanggal As String, strBulan As String, strTahun As String
    Dim strHari As String
    Dim intCol As Integer, intRow As Integer
    Dim blnAwalBulan As Boolean
    Dim firstMonthYear As Date
    Dim Datetgl As Date
    Dim iLoop As Integer

    blnAwalBulan = True
    firstMonthYear = CDate("01/" & str(Month(calDate)) & "/" & str(Year(calDate)))

    With Me.fgKalender
        For iLoop = 1 To nLastDay
            If blnAwalBulan Then
                Datetgl = firstMonthYear
            Else
                Datetgl = DateSerial(Format(Datetgl, "yyyy"), Format(Datetgl, "MM"), Val(Format(Datetgl, "dd")) + 1)
            End If
            strTanggal = Day(Datetgl)
            strBulan = Month(Datetgl)
            strHari = WeekdayName(Weekday(Datetgl), , vbSunday)

            If blnAwalBulan Then
                .Rows = .Rows + 1
                intRow = .Rows - 1
                blnAwalBulan = False
            End If
            Select Case strHari
                Case "Senin"
                    intCol = 0
                    intRow = .Rows - 1
                Case "Selasa"
                    intCol = 1
                Case "Rabu"
                    intCol = 2
                Case "Kamis"
                    intCol = 3
                Case "Jumat"
                    intCol = 4
                Case "Sabtu"
                    intCol = 5
                Case "Minggu"
                    intCol = 6
                    .Rows = .Rows + 1
            End Select
            .TextMatrix(intRow, intCol) = strTanggal
            If strTanggal = Day(Now) And strBulan = Month(Now) Then
                .row = intRow
                .Col = intCol
                .CellBackColor = &H4000&
                .CellForeColor = vbWhite
            End If
            If strHari = "Minggu" Or strHari = "Sabtu" Then
                .row = intRow
                .Col = intCol
                .CellBackColor = vbRed
                .CellFontBold = True
                .CellForeColor = vbWhite
            End If

            Dim ObjRs As New ADODB.recordset

            strSQL = "SELECT Tanggal, [Hari Libur] FROM v_tanggal" & _
            " WHERE DAY(Tanggal)='" & Day(Datetgl) & "'" & _
            " AND MONTH(Tanggal)='" & Month(Datetgl) & "'" & _
            " AND YEAR(Tanggal)='" & Year(Datetgl) & "'" & _
            " ORDER BY Tanggal"
            
            ObjRs.Open strSQL, dbConn, 3, 2
            If Not ObjRs.EOF Then
                If Not IsNull(ObjRs.Fields.Item("Hari Libur").Value) Then
                    .row = intRow
                    .Col = intCol
                    .CellBackColor = vbRed
                    .CellFontBold = True
                    .CellForeColor = vbWhite
                End If
            End If
            ObjRs.Close

        Next iLoop
    End With

End Sub

'24  April 2008
Private Sub subSetFGJadwalKerja()
    Dim intJumHari As Integer
    Dim i As Integer
    Dim itm As ListItem
    Dim strHari As String

    Call totalHari(Me.DTPicker1.Month, Me.DTPicker1.Year)
    intJumHari = TtlHari
    With Me.fgJadwalKerja
        .clear
        .Rows = 3
        .Cols = intJumHari + 1
        .FixedCols = 1
        .FixedRows = 2
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
        .TextMatrix(0, 0) = "Nama Pegawai"
        .TextMatrix(1, 0) = "Nama Pegawai"
        .ColWidth(0) = 3000
        .RowHeight(0) = 300
        For i = 1 To intJumHari
            Me.DTPicker1.Day = i
            strHari = WeekdayName(Weekday(Me.DTPicker1.Value), , vbSunday)

            .TextMatrix(0, i) = Format(Me.DTPicker1.Value, "MMMM yyyy")
            .TextMatrix(1, i) = i
            .row = 0
            .Col = i
            .CellAlignment = flexAlignCenterCenter
            .ColWidth(i) = 500
            .ColAlignment(i) = flexAlignCenterCenter
            If strHari = "Minggu" Or strHari = "Sabtu" Then
                .row = 1
                .CellBackColor = vbRed
                .CellForeColor = vbWhite
            Else
                .row = 1
                .CellBackColor = &H4000&
                .CellForeColor = vbWhite
            End If
            strSQL = "SELECT [Hari Libur] FROM v_tanggal" & _
            " WHERE DAY(Tanggal)='" & i & "'" & _
            " AND MONTH(Tanggal)='" & Month(Me.DTPicker1.Value) & "'" & _
            " AND YEAR(Tanggal)='" & Year(Me.DTPicker1.Value) & "'"
            Call msubRecFO(rs, strSQL)
            If Not rs.EOF Then
                If Not IsNull(rs.Fields.Item("Hari Libur").Value) Then
                    .row = 1
                    .CellBackColor = vbRed
                    .CellForeColor = vbWhite
                End If
            End If
        Next
        For Each itm In Me.lvjadwalkerja.ListItems
            .row = .Rows - 1
            .TextMatrix(.row, 0) = itm.Text
            .Rows = .Rows + 1
        Next
    End With
End Sub

'25      April 2008

Private Sub subIsiFgJadwalKerja()
    Dim intRowNama As Integer, intColTgl As Integer
    Dim strNama As String, strTanggal As String, strShift As String
    Dim strKodeShift As String
    Dim i As Integer

    strSQL = "SELECT * FROM V_JadwalKerjaNew" & _
    " WHERE Ruangan='" & dcTempatBertugas & "'" & _
    " AND MONTH(Tanggal)='" & Me.DTPicker1.Month & "'" & _
    " AND YEAR(Tanggal)='" & Me.DTPicker1.Year & "'" & _
    " ORDER BY Tanggal"

    Call msubRecFO(rs, strSQL)
    For i = 2 To fgJadwalKerja.Rows - 2
        For j = 1 To fgJadwalKerja.Cols - 1
            Me.fgJadwalKerja.TextMatrix(i, j) = "L"
        Next j
    Next i
    While Not rs.EOF
        strNama = rs.Fields.Item("Nama").Value
        strShift = rs.Fields.Item("Shift").Value
        strTanggal = rs.Fields.Item("Tanggal").Value

        intRowNama = funcCariRowNama(strNama)
        intColTgl = funcCariColTanggal(Day(strTanggal))

        strKodeShift = UCase$(Left(strShift, 1))
        Me.fgJadwalKerja.TextMatrix(intRowNama, intColTgl) = strKodeShift
        rs.MoveNext
    Wend
End Sub

Private Function funcCariRowNama(ByVal NamaPegawai As String) As Integer
    Dim intMaxRow As Integer
    Dim i As Integer

    With Me.fgJadwalKerja
        intMaxRow = .Rows - 1
        For i = 2 To intMaxRow
            If .TextMatrix(i, 0) = NamaPegawai Then
                funcCariRowNama = i
                Exit For
            Else
                funcCariRowNama = 0
            End If
        Next
    End With
End Function

Private Function funcCariColTanggal(ByVal tgl As String) As Integer
    Dim intMaxCol As Integer
    Dim i As Integer

    With Me.fgJadwalKerja
        intMaxCol = .Cols - 1
        For i = 1 To intMaxCol
            If .TextMatrix(1, i) = tgl Then
                funcCariColTanggal = i
                Exit For
            Else
                funcCariColTanggal = 0
            End If
        Next
    End With
End Function

'28      April 2008
Private Sub subSimpanJadwalBaru()
    On Error GoTo tangani

    Dim intMaxRow As Integer, intMaxCol As Integer
    Dim c As Integer, R As Integer
    Dim strKdTgl As String, strKdShift As String, strIDPegawai As String

    With Me.fgJadwalKerja
        intMaxRow = .Rows - 1
        intMaxCol = .Cols - 1

        For R = 2 To intMaxRow
            strIDPegawai = funcCariIdPegawai(.TextMatrix(R, 0))
            For c = 1 To intMaxCol
                strKdShift = funcCariKodeShift(.TextMatrix(R, c))
                Me.DTPicker1.Day = CInt(.TextMatrix(1, c))
                strSQL = "SELECT [Kode Tgl] FROM v_tanggal" & _
                " WHERE DAY(Tanggal)='" & Me.DTPicker1.Day & "'" & _
                " AND MONTH(Tanggal)='" & Me.DTPicker1.Month & "'" & _
                " AND YEAR(Tanggal)='" & Me.DTPicker1.Year & "'"
                Call msubRecFO(rs, strSQL)
                If Not rs.EOF Then
                    strKdTgl = rs.Fields.Item("Kode Tgl").Value
                End If
                If .TextMatrix(R, c) = "" Then
                    Call subDeleteJadwalDetailPerShift(strIDPegawai, Me.DTPicker1.Day & "/" & Me.DTPicker1.Month & "/" & Me.DTPicker1.Year)
                Else
                    Call subJadwalKerjaDetail(strIDPegawai, Me.DTPicker1.Day & "/" & Me.DTPicker1.Month & "/" & Me.DTPicker1.Year, strKdShift)
                End If
            Next
        Next
    End With
    MsgBox "Simpan jadwal selesai.", vbInformation, "Perhatian"
    Exit Sub

tangani:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Function funcCariIdPegawai(ByVal NamaPegawai As String) As String
    Dim itm As ListItem

    For Each itm In Me.lvjadwalkerja.ListItems
        If itm.Text = NamaPegawai Then
            funcCariIdPegawai = Right(itm.key, Len(itm.key) - 1)
            Exit For
        Else
            funcCariIdPegawai = ""
        End If
    Next
End Function

Private Function funcCariKodeShift(ByVal NamaShift As String) As String
    strSQL = "SELECT KdShift FROM ShiftKerja_New" & _
    " WHERE NamaShift LIKE '" & NamaShift & "%'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        funcCariKodeShift = rs.Fields.Item("KdShift").Value
    End If
End Function

Public Sub subJadwalKerja(ByVal idpegawai As String, ByVal KdShift As String)
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglJadwalKerja", adDate, adParamInput, , DTPicker1)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, idpegawai)
        .Parameters.Append .CreateParameter("KdShift", adChar, adParamInput, 2, KdShift)

        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AU_JADWALKERJA"
        .Execute

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Sub

Public Sub subJadwalKerjaDetail(ByVal idpegawai As String, ByVal tglJadwalKerja As Date, ByVal KdShift As String)
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, idpegawai)
'        .Parameters.Append .CreateParameter("JadwalKerja", adDate, adParamInput, , tglJadwalKerja)
        .Parameters.Append .CreateParameter("JadwalKerja", adDate, adParamInput, , CDate(tglJadwalKerja))
        .Parameters.Append .CreateParameter("KdShift", adChar, adParamInput, 2, KdShift)

        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AU_JADWALKERJADETAIL"
        .Execute

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Sub

Private Sub subTandaiTanggal(ByVal TanggalPilihan As String)
    Dim strTgl As String
    Dim i As Integer, j As Integer
    Dim intMaxRow As Integer, intMaxCol As Integer

    strTgl = Day(TanggalPilihan)
    With Me.fgJadwalKerja
        intMaxRow = .Rows - 1
        intMaxCol = .Cols - 1
        For i = 1 To intMaxCol
            .row = 1
            .Col = i
            If .CellBackColor = &H808080 Then
                For j = 1 To intMaxRow
                    .row = j
                    If .CellBackColor = vbYellow Then
                        .CellBackColor = &H808080
                    Else
                        .CellBackColor = vbWhite
                    End If
                Next
                .row = 1
                Me.DTPicker1.Day = CInt(.TextMatrix(.row, .Col))
                If WeekdayName(Weekday(Me.DTPicker1.Value), , vbSunday) = "Minggu" Then
                    .CellBackColor = vbRed
                Else
                    .CellBackColor = &H4000&
                End If
            End If
            If strTgl = .TextMatrix(1, i) Then
                For j = 1 To intMaxRow
                    .row = j
                    If .CellBackColor = &H808080 Then
                        .CellBackColor = vbYellow
                    Else
                        .CellBackColor = &H808080
                    End If
                Next
            End If
        Next
    End With
End Sub

Private Sub subDeleteJadwalPerShift(ByVal idpegawai As String, ByVal KodeTanggal As String)
    strSQL = "DELETE FROM JadwalKerja WHERE" & _
    " KdTgl='" & KodeTanggal & "'" & _
    " AND IdPegawai='" & idpegawai & "'"
    dbConn.Execute strSQL
End Sub

Private Sub subDeleteJadwalDetailPerShift(ByVal idpegawai As String, ByVal tglJadwalKerja As Date)
    strSQL = "DELETE FROM RealisasiJadwalKerja WHERE" & _
    " IdPegawai='" & idpegawai & "'" & _
    " AND Hari='" & Format(tglJadwalKerja, "dd") & "'" & _
    " AND Bulan='" & Format(tglJadwalKerja, "MM") & "'" & _
    " AND Tahun='" & Format(tglJadwalKerja, "yyyy") & "'"
    dbConn.Execute strSQL
End Sub

Private Sub txtNamaPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

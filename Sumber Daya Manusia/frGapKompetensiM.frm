VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGapKompetensiM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Gap Kompetensi"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frGapKompetensiM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   9225
   Begin MSComctlLib.TreeView tv 
      Height          =   2415
      Left            =   360
      TabIndex        =   15
      Top             =   5160
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4260
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
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
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Pangkat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   9015
      Begin VB.CommandButton cmdMin 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   13
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton cmdPlus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   12
         Top             =   1320
         Width           =   255
      End
      Begin VB.ListBox lstKebutuhan 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   10800
         TabIndex        =   10
         Top             =   1920
         Width           =   6015
      End
      Begin VB.TextBox txtSkill 
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
         ForeColor       =   &H80000007&
         Height          =   810
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   6840
         Width           =   7815
      End
      Begin MSComctlLib.ListView lvPendidikan 
         Height          =   2175
         Left            =   11040
         TabIndex        =   7
         Top             =   4200
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Kualifikasi Pendidikan"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcJabatan 
         Height          =   390
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView tvPendidikan 
         Height          =   2415
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4260
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Kebutuhan Diklat"
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
         Left            =   240
         TabIndex        =   11
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Skill"
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
         Left            =   240
         TabIndex        =   9
         Top             =   6600
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Standarisasi Pendidikan"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1680
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nama Jabatan"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   9000
      Width           =   1215
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
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
      Picture         =   "frGapKompetensiM.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frGapKompetensiM.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14895
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4800
      Picture         =   "frGapKompetensiM.frx":4CE9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmGapKompetensiM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
    dcJabatan.Text = ""
    For i = 1 To lvPendidikan.ListItems.Count
        lvPendidikan.ListItems(i).Checked = False
    Next
    lstKebutuhan.clear
End Sub

Private Sub cmdMin_Click()
    If MsgBox("Hapus ?", vbYesNo, "..:.") = vbNo Then Exit Sub
    lstKebutuhan.RemoveItem lstKebutuhan.ListIndex
End Sub

Private Sub cmdPlus_Click()
Dim s As String
    s = InputBox("Kebutuhan diklat", "..:.")
    
    lstKebutuhan.AddItem s, lstKebutuhan.ListCount
End Sub

Private Sub cmdSimpan_Click()
    strSQL = "select * from GapDiklat where kdJabatan='" & dcJabatan.BoundText & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount = 0 Then
        strSQL = "insert into GapDiklat values ('" & dcJabatan.BoundText & "','" & txtSkill.Text & "')"
        Call msubRecFO(rs, strSQL)
        For i = 1 To tvPendidikan.Nodes.Count
            If tvPendidikan.Nodes(i).Checked = True Then
                strSQL = "insert into GapDiklat_K values ('" & dcJabatan.BoundText & "','" & Right(lvPendidikan.ListItems(i).Key, 4) & "')"
                Call msubRecFO(rs, strSQL)
            End If
        Next
        For i = 1 To tv.Nodes.Count
            If tv.Nodes(i).Checked = True Then
                strSQL = "insert into GapDiklat_KD values ('" & dcJabatan.BoundText & "','" & lstKebutuhan.List(i) & "')"
                Call msubRecFO(rs, strSQL)
            End If
        Next
    Else
        strSQL = "delete from GapDiklat where kdjabatan ='" & dcJabatan.BoundText & "'"
        Call msubRecFO(rs, strSQL)
        strSQL = "insert into GapDiklat values ('" & dcJabatan.BoundText & "','" & txtSkill.Text & "')"
        Call msubRecFO(rs, strSQL)
        
        strSQL = "delete from GapDiklat_K where kdjabatan ='" & dcJabatan.BoundText & "'"
        Call msubRecFO(rs, strSQL)
        For i = 1 To tvPendidikan.Nodes.Count
            If tvPendidikan.Nodes(i).Checked = True Then
                strSQL = "insert into GapDiklat_K values ('" & dcJabatan.BoundText & "','" & Right(tvPendidikan.Nodes(i).Key, 4) & "')"
                Call msubRecFO(rs, strSQL)
            End If
        Next
        
        strSQL = "delete from GapDiklat_KD where kdjabatan ='" & dcJabatan.BoundText & "'"
        Call msubRecFO(rs, strSQL)
        For i = 1 To tv.Nodes.Count
            If tv.Nodes(i).Checked = True Then
                strSQL = "insert into GapDiklat_KD values ('" & dcJabatan.BoundText & "','" & Right(tv.Nodes(i).Key, 5) & "')"
                Call msubRecFO(rs, strSQL)
            End If
        Next
    End If
    
    MsgBox "Data berhasil Disimpan .?", vbInformation, "..:."
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub LoadGapDiklat()
Dim ii As Integer
    
    strSQL = " select * from GapDiklat where kdJabatan='" & dcJabatan.BoundText & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then txtSkill.Text = rs(1)
    
    For ii = 1 To tvPendidikan.Nodes.Count
        tvPendidikan.Nodes(ii).Checked = False
        If tvPendidikan.Nodes(ii).Children = 0 Then
            tvPendidikan.Nodes(ii).Parent.Expanded = False
        End If
    Next
    strSQL = " select * from GapDiklat_k where kdJabatan='" & dcJabatan.BoundText & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        For i = 0 To rs.RecordCount - 1
            For ii = 1 To tvPendidikan.Nodes.Count
                If Right(tvPendidikan.Nodes(ii).Key, 4) = rs(1) Then
                    tvPendidikan.Nodes(ii).Checked = True
                    tvPendidikan.Nodes(ii).Parent.Expanded = True
                End If
            Next
            rs.MoveNext
        Next
    End If
    
    For ii = 1 To tv.Nodes.Count
        tv.Nodes(ii).Checked = False
        If tv.Nodes(ii).Children = 0 Then tv.Nodes(ii).Parent.Expanded = False
    Next
    strSQL = " select * from GapDiklat_kd where kdJabatan='" & dcJabatan.BoundText & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        For i = 0 To rs.RecordCount - 1
            For ii = 1 To tv.Nodes.Count
                If Right(tv.Nodes(ii).Key, 5) = rs(1) Then
                    tv.Nodes(ii).Checked = True
                    tv.Nodes(ii).Parent.Expanded = True
                End If
            Next
            rs.MoveNext
        Next
    End If
End Sub

Private Sub dcJabatan_Change()
    Call LoadGapDiklat
End Sub

Private Sub Form_Load()
     Call centerForm(Me, MDIUtama)
     Call PlayFlashMovie(Me)
     Call subLoadDcSource
'     Call GridSource
End Sub

Private Sub subLoadDcSource()
'    strSQL = "select * from Jabatan where StatusEnabled='1' order by NamaJabatan "
'    Call msubRecFO(rs, strSQL)
'    Set dcJabatan.RowSource = rs
'    dcJabatan.BoundColumn = rs(0)
'    dcJabatan.ListField = rs(1)
    
    Call msubDcSource(dcJabatan, rs, "select * from Jabatan where StatusEnabled='1' order by NamaJabatan ")
'    Call msubDcSource(lvPendidikan, rs, "select * from KualifikasiJurusan where StatusEnabled='1' order by KualifikasiJurusan ")
    
'    'lvPendidikan
    strSQL = "select * from KualifikasiJurusan where StatusEnabled='1' order by KualifikasiJurusan "
    Call msubRecFO(rs, strSQL)
    For i = 0 To rs.RecordCount - 1
        lvPendidikan.ListItems.Add , "A" & rs(0), rs(1)
        rs.MoveNext
    Next
    
    Dim txtKey0, txtKey1 As String
    Dim nodX As Node
    
    strSQL = "select distinct KdJenisDiklat ,JenisDiklat  from V_JenisDiklat"
    Call msubRecFO(rs, strSQL)
    For i = 0 To rs.RecordCount - 1
        txtKey0 = "A~" & rs(0)
        Set nodX = tv.Nodes.Add(, , txtKey0, rs(1))
        rs.MoveNext
    Next
    strSQL = "select * from V_JenisDiklat "
    Call msubRecFO(rs, strSQL)
    For i = 0 To rs.RecordCount - 1
        txtKey0 = "A~" & rs(0)
        txtKey1 = "B~" & rs(2)
        Set nodX = tv.Nodes.Add(txtKey0, tvwChild, txtKey1, rs(3))
        rs.MoveNext
    Next
    
    strSQL = "SELECT distinct KdPendidikan, Pendidikan FROM V_JenisPendidikan "
    Call msubRecFO(rs, strSQL)
    For i = 0 To rs.RecordCount - 1
        txtKey0 = "A~" & rs(0)
        Set nodX = tvPendidikan.Nodes.Add(, , txtKey0, rs(1))
        rs.MoveNext
    Next
    strSQL = "SELECT KdPendidikan, Pendidikan, KdKualifikasiJurusan, KualifikasiJurusan FROM V_JenisPendidikan order by NoUrut"
    Call msubRecFO(rs, strSQL)
    For i = 0 To rs.RecordCount - 1
        txtKey0 = "A~" & rs(0)
        txtKey1 = "B~" & rs(2)
        Set nodX = tvPendidikan.Nodes.Add(txtKey0, tvwChild, txtKey1, rs(3))
        rs.MoveNext
    Next
    
'    lvPendidikan.BoundColumn = rs(0)
'    lvPendidikan.ListField = rs(1)
End Sub


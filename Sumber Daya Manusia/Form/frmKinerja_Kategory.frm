VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKinerja_Kategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Master Kinerja Pegawai"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKinerja_Kategory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11250
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   9960
      TabIndex        =   4
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Height          =   6855
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   11055
      Begin TabDlg.SSTab SSTab2 
         Height          =   6255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   11033
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Master Kinerja"
         TabPicture(0)   =   "frmKinerja_Kategory.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label22"
         Tab(0).Control(1)=   "Label23"
         Tab(0).Control(2)=   "Label24"
         Tab(0).Control(3)=   "Label25"
         Tab(0).Control(4)=   "dgData"
         Tab(0).Control(5)=   "dcKategoryKinerja"
         Tab(0).Control(6)=   "txtKdKinerja"
         Tab(0).Control(7)=   "txtNmKinerja"
         Tab(0).Control(8)=   "Check1"
         Tab(0).Control(9)=   "Text3"
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Kinerja Pegawai"
         TabPicture(1)   =   "frmKinerja_Kategory.frx":0CE6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label26"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "tvDiagnosa"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Text4"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Text5"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lstPegawai"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin VB.ListBox lstPegawai 
            Appearance      =   0  'Flat
            Height          =   3600
            Left            =   3120
            TabIndex        =   22
            Top             =   960
            Visible         =   0   'False
            Width           =   7455
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3120
            TabIndex        =   20
            Top             =   600
            Width           =   7455
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   18
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   -73920
            MaxLength       =   30
            TabIndex        =   14
            Top             =   5760
            Width           =   9495
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   -65880
            TabIndex        =   13
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNmKinerja 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   -73560
            TabIndex        =   11
            Top             =   1800
            Width           =   5895
         End
         Begin VB.TextBox txtKdKinerja 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   330
            Left            =   -73560
            MaxLength       =   3
            TabIndex        =   9
            Top             =   840
            Width           =   1320
         End
         Begin MSDataListLib.DataCombo dcKategoryKinerja 
            Height          =   330
            Left            =   -73560
            TabIndex        =   8
            Top             =   1320
            Width           =   1320
            _ExtentX        =   2328
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
         Begin MSDataGridLib.DataGrid dgData 
            Height          =   3450
            Left            =   -74880
            TabIndex        =   15
            Top             =   2280
            Width           =   10440
            _ExtentX        =   18415
            _ExtentY        =   6085
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
         Begin MSComctlLib.TreeView tvDiagnosa 
            Height          =   5055
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   8916
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
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Pegawai"
            Height          =   210
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Kategory"
            Height          =   210
            Left            =   -74760
            TabIndex        =   17
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "Cari Kinerja"
            Height          =   255
            Left            =   -74880
            TabIndex        =   16
            Top             =   5805
            Width           =   1455
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Materi Kinerja"
            Height          =   210
            Left            =   -74760
            TabIndex        =   12
            Top             =   1740
            Width           =   1095
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kinerja"
            Height          =   210
            Left            =   -74760
            TabIndex        =   10
            Top             =   885
            Width           =   1020
         End
      End
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   8040
      Width           =   1215
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
   Begin VB.Image Image4 
      Height          =   945
      Left            =   9360
      Picture         =   "frmKinerja_Kategory.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKinerja_Kategory.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmKinerja_Kategory.frx":30E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmKinerja_Kategory.frx":5AA9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmKinerja_Kategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub cmdCancel_Click()
Dim i, ii As Integer
Dim kode() As String

    Text4.Text = ""
    Text5.Text = ""
        
    For ii = 1 To tvDiagnosa.Nodes.Count
        tvDiagnosa.Nodes(ii).Checked = False
    Next
    
    txtKdKinerja.Text = ""
    txtNmKinerja.Text = ""
    dcKategoryKinerja.Text = ""
    Check1.Value = Checked
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim i, ii As Integer
Dim kode() As String
Dim sttsebel As String
Dim kode2 As String

    If Check1.Value = Checked Then
        sttsebel = "1"
    Else
        sttsebel = "0"
    End If
    
    If SSTab2.Tab = 0 Then
        If txtKdKinerja.Text = "" Then 'ADD NEW
            If txtNmKinerja.Text = "" Then Exit Sub
            strSQL = "select max(KdKinerja) from MasterKinerja "
            Call msubRecFO(rs, strSQL)
            If rs.RecordCount = 0 Then
                kode2 = "01"
            Else
                kode2 = Format(Val(rs(0)) + 1, "0##")
            End If
            
            strSQL = "insert into MasterKinerja values ('" & kode2 & "','" & dcKategoryKinerja.BoundText & "','" & txtNmKinerja.Text & "','" & sttsebel & "')"
            Call msubRecFO(rs, strSQL)
        Else 'UPDATE
            If txtNmKinerja.Text = "" Then Exit Sub
            strSQL = "update MasterKinerja set NamaKinerja ='" & txtNmKinerja.Text & "', KdKategoryKinerja='" & dcKategoryKinerja.BoundText & "',StatusEnabled ='" & sttsebel & "' where KdKinerja='" & txtKdKinerja.Text & "'"
            Call msubRecFO(rs, strSQL)
        End If
        Call LoadDataGrid
    End If
    If SSTab2.Tab = 1 Then
        If Text5.Text = "" Then Exit Sub
        strSQL = "delete from KinerjaPegawai where idpegawai='" & Text4.Text & "'"
        Call msubRecFO(rs, strSQL)
        If SSTab2.Tab = 1 Then
            For ii = 1 To tvDiagnosa.Nodes.Count
                If tvDiagnosa.Nodes(ii).Checked = True Then
                    kode = Split(tvDiagnosa.Nodes(ii).key, "~")
                    If kode(0) = "B" Then
                        strSQL = "insert into KinerjaPegawai values ('" & kode(1) & "','" & Text4.Text & "')"
                        Call msubRecFO(rs, strSQL)
                    End If
                End If
            Next
        End If
    End If
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdKinerja.Text = dgData.Columns(0)
    dcKategoryKinerja.Text = dgData.Columns(1)
    txtNmKinerja.Text = dgData.Columns(2)
    If dgData.Columns(3) = "1" Then
        Check1.Value = Checked
    Else
        Check1.Value = Unchecked
    End If
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call LoadDiagnosa2
    Call LoadDataGrid
    SSTab2.Tab = 0
End Sub

Private Sub LoadDataGrid()
Call msubDcSource(dcKategoryKinerja, rs, "select * from KategoryKinerja  WHERE StatusEnabled=1")
    strSQL = "SELECT     MasterKinerja.KdKinerja, KategoryKinerja.KategoryKinerja, MasterKinerja.NamaKinerja, MasterKinerja.StatusEnabled,MasterKinerja.KdKategoryKinerja " & _
             "FROM         MasterKinerja INNER JOIN KategoryKinerja ON MasterKinerja.KdKategoryKinerja = KategoryKinerja.KdKategoryKinerja " & _
             " where MasterKinerja.NamaKinerja like '%" & Text3.Text & "%'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        Set dgData.DataSource = rs
        
        dgData.Columns(0).Caption = "Kode Kinerja"
        dgData.Columns(1).Caption = "Kategory"
        dgData.Columns(2).Caption = "Nama Kinerja"
        dgData.Columns(3).Caption = "Enabled"
        dgData.Columns(0).Width = 700
        dgData.Columns(1).Width = 1000
        dgData.Columns(2).Width = 7000
        dgData.Columns(3).Width = 1000
        dgData.Columns(4).Width = 0
    End If
End Sub


Private Function sp_aaa_KategoryKinerja() As Boolean
    On Error GoTo errLoad

    sp_aaa_KategoryKinerja = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adChar, adParamInput, 9, txtKdBarang.Text)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dcAsalBarang.BoundText)
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, txtNoTerima.Text)
        
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, S_kdruangan)
        .Parameters.Append .CreateParameter("JmlMin", adDouble, adParamInput, , S_JmlMin)
        .Parameters.Append .CreateParameter("JmlMax", adDouble, adParamInput, , S_JmlMax)
        
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "U")

        .ActiveConnection = dbConn
        .CommandText = "AUD_KategoryKinerja"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
            sp_aaa_KategoryKinerja = False
        Else
'            txtNoClosing.Text = .Parameters("OutputNoClosing").Value
        End If
    End With

    Exit Function
errLoad:
    sp_aaa_KategoryKinerja = False
'    Resume 0
    Call msubPesanError
End Function


Private Sub LoadDiagnosa2()
Dim txtKey0, txtKey1, txtKey2, txtKey3, txtKey4, txtKey5, txtKey6, txtKey7, txtNama, txtSatuan As String
Dim nodX As Node

'    If dcKeluhan.MatchedWithList = False Then Exit Sub
    tvDiagnosa.Nodes.clear
    
    Set nodX = tvDiagnosa.Nodes.add(, , "A~01", "BHU")
    Set nodX = tvDiagnosa.Nodes.add(, , "A~02", "BPU")
    
    strSQL = "select * from MasterKinerja where statusEnabled='1'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then Exit Sub
    Do
'        If UCase(rs(2)) = "TEKS" Then
'            txtNama = rs(1) & " : . . . "
'            If rs(3) <> "" Then
'                txtNama = txtNama & rs(3)
'            End If
'        Else
            txtNama = rs(2)
'        End If
        
        Select Case Val(rs(1))
            Case 1
                txtKey1 = "B~" & rs(0)
                'Set nodX = tvDiagnosa.Nodes.add(, , txtKey0, txtNama)
                Set nodX = tvDiagnosa.Nodes.add("A~01", tvwChild, txtKey1, txtNama)
            Case 2
                txtKey1 = "B~" & rs(0)
                Set nodX = tvDiagnosa.Nodes.add("A~02", tvwChild, txtKey1, txtNama)
                
'            Case 2
'                txtKey2 = "C~" & rs(0)
'                Set nodX = tvDiagnosa.Nodes.add(txtKey1, tvwChild, txtKey2, txtNama)
'
'            Case 3
'                txtKey3 = "D~" & rs(0)
'                Set nodX = tvDiagnosa.Nodes.add(txtKey2, tvwChild, txtKey3, txtNama)
'
'            Case 4
'                txtKey4 = "E~" & rs(0)
'                Set nodX = tvDiagnosa.Nodes.add(txtKey3, tvwChild, txtKey4, txtNama)
'
'            Case 5
'                txtKey5 = "F~" & rs(0)
'                Set nodX = tvDiagnosa.Nodes.add(txtKey4, tvwChild, txtKey5, txtNama)
'
'            Case 6
'                txtKey6 = "G~" & rs(0)
'                Set nodX = tvDiagnosa.Nodes.add(txtKey5, tvwChild, txtKey6, txtNama)
'
'            Case 7
'                txtKey7 = "H~" & rs(0)
'                Set nodX = tvDiagnosa.Nodes.add(txtKey6, tvwChild, txtKey7, txtNama)
                
        End Select
        
'        Set nodx = tvDiagnosa.Nodes.Add(, , "A~" & rs(0), rs(1))
        rs.MoveNext
    Loop Until rs.EOF

End Sub


Private Sub lstPegawai_DblClick()
    lstPegawai.Visible = False
    strSQL = "select * from datapegawai where namalengkap ='" & lstPegawai.List(lstPegawai.ListIndex) & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        Text5.Text = rs(3)
        Text4.Text = rs(0)
        Text5.SetFocus
        Call Text4_KeyDown(13, False)
    End If
End Sub

Private Sub lstPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call lstPegawai_DblClick
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
    If SSTab2.Tab = 0 Then
        Call LoadDataGrid
    End If
    If SSTab2.Tab = 1 Then
        Call LoadDiagnosa2
    End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call LoadDataGrid
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim i, ii As Integer
        Dim kode() As String
        
        
        strSQL = "select * from DataPegawai where idpegawai='" & Text4.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            Text4.Text = rs(0)
            Text5.Text = rs(3)
        End If
        
        For ii = 1 To tvDiagnosa.Nodes.Count
            tvDiagnosa.Nodes(ii).Checked = False
        Next
        strSQL = "select * from KinerjaPegawai where idpegawai='" & Text4.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            For i = 0 To rs.RecordCount - 1
                For ii = 1 To tvDiagnosa.Nodes.Count
                    kode = Split(tvDiagnosa.Nodes(ii).key, "~")
                    If kode(0) = "B" And kode(1) = rs(0) Then
                        tvDiagnosa.Nodes(ii).Checked = True
                        Exit For
                    End If
                Next
                rs.MoveNext
            Next
        End If
    End If
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 And lstPegawai.Visible = True Then lstPegawai.SetFocus
    If KeyCode = 13 Then
        If Text5.Text <> "" Then
        strSQL = "select * from datapegawai where NamaLengkap like '%" & Text5.Text & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            lstPegawai.Visible = True
            lstPegawai.clear
            For i = 0 To rs.RecordCount - 1
                lstPegawai.AddItem rs(3), i
                rs.MoveNext
            Next
            
        End If
    End If
    End If
End Sub


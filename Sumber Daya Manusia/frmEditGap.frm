VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEditGap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medifirst2000 - Edit Gap"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12810
   Icon            =   "frmEditGap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   12735
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11430
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
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
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   12735
      Begin VB.TextBox txtJabatan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   3315
      End
      Begin VB.TextBox txtIdPegawai 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox txtNamaPegawai 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   3315
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   3135
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   5530
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorSel    =   -2147483643
         BackColorBkg    =   -2147483633
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
      Begin VB.Label lblInstansi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label lblInstansi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   285
         Width           =   915
      End
      Begin VB.Label lblInstansi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   645
         Width           =   1185
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmEditGap.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmEditGap.frx":2328
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   10920
      Picture         =   "frmEditGap.frx":4CE9
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
   End
End
Attribute VB_Name = "frmEditGap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    
    Call SETGRID
    Call LOADDATA
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call LOADDATA
    End If
End Sub

Private Sub LOADDATA()
    If txtIdPegawai.Text = "" Then Exit Sub
    
    strSQL = "SELECT * FROM V_ListPegawai where idpegawai='" & txtIdPegawai.Text & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        txtNamaPegawai.Text = rs(1)
        txtJabatan.Text = rs(6)
        
        
    End If
    
End Sub

Private Sub SETGRID()
    With fgData
        .Cols = 10
        .Rows = 3
        
        .TextMatrix(0, 1) = "Jabatan"
        .TextMatrix(1, 1) = "Jabatan"
        .MergeCells = flexMergeFree
        .MergeCol(1) = True
        
        .TextMatrix(0, 2) = "Job Kualifikasi"
        .TextMatrix(0, 3) = "Job Kualifikasi"
        .TextMatrix(0, 4) = "Job Kualifikasi"
        .MergeCells = flexMergeFree
        .MergeRow(0) = True

        .TextMatrix(1, 2) = "Standarisasi Pendidikan"
        .TextMatrix(1, 3) = "Pendidikan Real"
        .TextMatrix(1, 4) = "GAP %"
        
        .TextMatrix(0, 5) = "Skill"
        .TextMatrix(1, 5) = "Skill"
        .MergeCells = flexMergeFree
        .MergeCol(5) = True
        
        .TextMatrix(0, 6) = "Diklat"
        .TextMatrix(0, 7) = "Diklat"
        .TextMatrix(0, 8) = "Diklat"
        .TextMatrix(0, 9) = "Diklat"
        
        .TextMatrix(1, 6) = "Kebutuhan"
        .TextMatrix(1, 7) = "Real (Training)"
        .TextMatrix(1, 8) = "Real (Training)"
        .TextMatrix(1, 9) = "GAP%"
        .MergeCells = flexMergeFree
        .MergeRow(1) = True
        
        .ColWidth(0) = 300
        .ColWidth(1) = 2000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 500
        .ColWidth(5) = 2000
        .ColWidth(6) = 3000
        .ColWidth(7) = 500
        .ColWidth(8) = 500
        .ColWidth(9) = 500
        
        Dim ii As Integer
        Dim i As Integer
        For i = 0 To 1
            For ii = 0 To 9
                .Row = i
                .Col = ii
                .CellBackColor = &H8000000F
            Next
        Next
    End With
End Sub


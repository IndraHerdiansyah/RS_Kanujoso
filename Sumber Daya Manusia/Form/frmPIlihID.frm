VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmPIlihID 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pilih ID untuk PIN"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmPIlihID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6135
   Begin VB.CommandButton cmdPlih 
      Caption         =   "&Pilih"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5895
      Begin MSDataGridLib.DataGrid dgIdUnPIN 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4471
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
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
   Begin VB.Label lblJumlahData 
      AutoSize        =   -1  'True
      Caption         =   "<data>"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4320
      Picture         =   "frmPIlihID.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPIlihID.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmPIlihID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strIdPilih As String

Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdPlih_Click()
    Dim sql As String
    sql = "SELECT * FROM v_PIN WHERE ID=" & funcPrepareString(strIdPilih)
    Set rs = New ADODB.recordset
    rs.Open sql, dbConn, adOpenForwardOnly, adLockReadOnly

    With frmPINAbsensiPegawai.ListView1.ListItems.Item(idxListViewPIN)
        .SubItems(2) = "T"
        .SubItems(3) = IIf(IsNull(rs.Fields.Item("Nama").Value), "", rs.Fields.Item("Nama").Value)
        .SubItems(4) = IIf(IsNull(rs.Fields.Item("JK").Value), "", rs.Fields.Item("JK").Value)
        .SubItems(5) = IIf(IsNull(rs.Fields.Item("ID").Value), "", rs.Fields.Item("ID").Value)
        .SubItems(6) = IIf(IsNull(rs.Fields.Item("Ruangan").Value), "", rs.Fields.Item("Ruangan").Value)
        .SubItems(7) = IIf(IsNull(rs.Fields.Item("Jabatan").Value), "", rs.Fields.Item("Jabatan").Value)
        .SubItems(8) = IIf(IsNull(rs.Fields.Item("Tgl. Daftar").Value), "", .SubItems(7) = rs.Fields.Item("Tgl. Daftar").Value)
    End With

    rs.Close
    Unload Me
End Sub

Private Sub dgIdUnPIN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    strIdPilih = Me.dgIdUnPIN.Columns("ID").Value
End Sub

Private Sub Form_Load()

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    Dim sql As String
    sql = "SELECT ID, Nama, JK, Ruangan, Jabatan, [Tgl. Daftar] FROM v_PIN" & _
    " WHERE (PIN IS NULL)"
    Set rs = New ADODB.recordset
    rs.Open sql, dbConn, adOpenForwardOnly, adLockReadOnly
    Set Me.dgIdUnPIN.DataSource = rs
    Me.lblJumlahData.Caption = rs.RecordCount & " data"
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJobList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Order List"
   ClientHeight    =   6465
   ClientLeft      =   390
   ClientTop       =   1260
   ClientWidth     =   14235
   Icon            =   "frmJobList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   14235
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   14055
      Begin VB.CommandButton cmdJobOrder 
         Caption         =   "Permintaan Baru"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtKdTask 
         Height          =   285
         Left            =   11280
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdRiwayat 
         Caption         =   "Riwayat"
         Height          =   255
         Left            =   11520
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkSelesai 
         Caption         =   "Status Job sudah Selesai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Klik 2x Untuk Melihat Riwayat Permintaan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
   End
   Begin MSComctlLib.ListView lvdata 
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No Permintaan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pengirim"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Masalah"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "TglOrder"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Penerima"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tanggal dikerjakan"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tanggal Selesai"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Prioritas"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.CommandButton cmdPerbaruiData 
      Caption         =   "Perbarui Data"
      Height          =   495
      Left            =   8760
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker dtAwal 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy "
      Format          =   127598595
      UpDown          =   -1  'True
      CurrentDate     =   40321
   End
   Begin MSComCtl2.DTPicker dtAkhir 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy "
      Format          =   127598595
      UpDown          =   -1  'True
      CurrentDate     =   40321
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal Order"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frmJobList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lv As ListItem
Dim i As Integer

Private Sub ambildata()
    Set rs = Nothing
    lvdata.ListItems.clear
    If chkSelesai.Value = 0 Then
        strSQL = "Select KdTask, Pengirim, Masalah, TglOrder,  PenanggungJawab, Status, TglMulai, TglSelesai, Prioritas  from V_SimOrder where  KdStatus in ('01', '06') And KdRuangan = '" & mstrKdRuangan & "' And TglOrder BETWEEN '" & Format(dtAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtAkhir.Value, "yyyy/MM/dd 23:59:59") & "' Order by TglOrder"

    ElseIf chkSelesai.Value = 1 Then
        strSQL = "Select  KdTask, Pengirim, Masalah, TglOrder,  PenanggungJawab, Status, TglMulai, TglSelesai, Prioritas from V_SimOrder where  KdStatus = '03' And TglOrder BETWEEN '" & Format(dtAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtAkhir.Value, "yyyy/MM/dd 23:59:59") & "' And KdRuangan = '" & mstrKdRuangan & "' Order by TglOrder"

    End If
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Or rs.BOF = True Then
        Exit Sub
    End If
    rs.MoveFirst
    With lvdata
        For i = 1 To rs.RecordCount
            Set x = lvdata.ListItems.add(, , rs.Fields(0).Value)
            x.SubItems(1) = rs.Fields(1).Value
            x.SubItems(2) = rs.Fields(2).Value
            x.SubItems(3) = rs.Fields(3).Value
            x.SubItems(4) = rs.Fields(4).Value
            x.SubItems(5) = rs.Fields(5).Value
            x.SubItems(6) = rs.Fields(6).Value
            x.SubItems(7) = rs.Fields(7).Value
            x.SubItems(8) = rs.Fields(8).Value
            rs.MoveNext
        Next i
    End With

End Sub

Private Sub chkSelesai_Click()
    ambildata
End Sub

Private Sub cmdJobOrder_Click()
    Unload frmJobList
    FrmTaskList.Show
End Sub

Private Sub cmdPerbaruiData_Click()
    ambildata
End Sub

Private Sub dcRuangan_Click(Area As Integer)
    ambildata
End Sub

Private Sub cmdRiwayat_Click()
    strNoRequest = Trim(txtKdTask.Text)
    frmDetailRequest.Show
End Sub

Private Sub Form_Load()
    Call centerForm(frmJobList, MDIUtama)
    dtAwal.Value = Now
    dtAkhir.Value = Now
    ambildata

End Sub

Private Sub lvdata_Click()
    On Error Resume Next
    txtKdTask.Text = lvdata.SelectedItem.Text
End Sub

Private Sub lvdata_DblClick()
    Call cmdRiwayat_Click
End Sub

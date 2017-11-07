VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPercobaan 
   Caption         =   "Form Percobaan"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   13845
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9840
      TabIndex        =   23
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Bandingkan"
      Height          =   375
      Left            =   11640
      TabIndex        =   22
      Top             =   1320
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   9840
      TabIndex        =   20
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "HH:mm"
      Format          =   478937091
      UpDown          =   -1  'True
      CurrentDate     =   39990
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   9495
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   495
         Left            =   7080
         TabIndex        =   19
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   6000
         TabIndex        =   18
         Top             =   3960
         Width           =   975
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   4200
         Top             =   4920
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"frmPercobaan.frx":0000
         OLEDBString     =   $"frmPercobaan.frx":0096
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "datatanggal"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4560
         TabIndex        =   17
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   478937089
         CurrentDate     =   39553
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   7080
         TabIndex        =   16
         Top             =   3960
         Width           =   1455
      End
      Begin MSComCtl2.MonthView MonthView1 
         Bindings        =   "frmPercobaan.frx":012C
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   3
         EndProperty
         Height          =   2310
         Left            =   4560
         TabIndex        =   15
         Top             =   1440
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         OLEDropMode     =   1
         StartOfWeek     =   478937090
         CurrentDate     =   39553
      End
      Begin VB.TextBox txtDisplay 
         Height          =   4335
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   3375
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtAlamatFRS 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   8280
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ListBox lstPIN 
      Height          =   1620
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdCek 
      Caption         =   "Cek Data"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtHasil 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton cmdFullUpload 
      Caption         =   "Full Upload"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtFRS 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid dgPIN 
      Height          =   2655
      Left            =   4560
      TabIndex        =   9
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4683
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
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   9840
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "HH:mm"
      Format          =   478937091
      UpDown          =   -1  'True
      CurrentDate     =   39990
   End
   Begin VB.Label lblJumlah 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "FRS"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmPercobaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCek_Click()
    Dim hasil As String
    hasil = cekSum(Me.txtHasil.Text)
    MsgBox hasil
End Sub

Private Sub cmdFullUpload_Click()
    subCekJumlahPIN Me.txtFRS.Text
    frmAbsensiPegawai.minta_absensi = False
End Sub

Private Sub cmdSimpan_Click()
    Set adoComm = New ADODB.Command
    adoComm.ActiveConnection = dbConn
    adoComm.CommandText = "UPDATE PINAbsensiPegawai SET " & _
    "AlamatFRS=" & funcPrepareString(Me.txtAlamatFRS.Text) & " " & _
    "WHERE IdPegawai=" & funcPrepareString(Me.txtID.Text)
    adoComm.CommandType = adCmdText
    adoComm.Execute
    Form_Load
End Sub

Private Sub Command2_Click()
    Dim t As Date

    strSQL = "SELECT Tanggal FROM v_tanggal"
    Call msubRecFO(rs, strSQL)
    While Not rs.EOF
        t = rs.Fields.Item("Tanggal").Value
        Me.MonthView1.Value = t
        Me.MonthView1.DayBold(t) = True
        rs.MoveNext
    Wend
    Me.MonthView1.Value = Now
End Sub

Private Sub Command3_Click()
    MsgBox Me.DTPicker1.DayOfWeek
End Sub

Private Sub Command4_Click()
    If Format(Me.DTPicker2.Value, "HH:mm") > Format(Me.DTPicker3.Value, "HH:mm") Then
        Me.Text2.Text = "Telat!!"
    ElseIf Format(Me.DTPicker2.Value, "HH:mm") < Format(Me.DTPicker3.Value, "HH:mm") Then
        Me.Text2.Text = "Ga telat"
    Else
        Me.Text2.Text = "Tepat waktu coy"
    End If
End Sub

Private Sub dgPIN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    Me.txtID.Text = Me.dgPIN.Columns("ID").Value
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    strSQL = "SELECT * FROM v_PIN"
    dbConn.Execute strSQL
    Call msubRecFO(rs, strSQL)
    Set dgPIN.DataSource = rs
    dgPIN.Columns("Tgl. Mulai").Width = 0

    Dim t As Date

    strSQL = "SELECT Tanggal FROM v_tanggal"
    Call msubRecFO(rs, strSQL)
    While Not rs.EOF
        t = rs.Fields.Item("Tanggal").Value
        Me.MonthView1.Value = t
        Me.MonthView1.DayBold(t) = True
        rs.MoveNext
    Wend
    Me.MonthView1.Value = Now
    Exit Sub
errLoad:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    MsgBox DateClicked & vbCrLf & WeekdayName(Weekday(DateClicked), , vbSunday) & _
    vbCrLf & Weekday(DateClicked) & vbCrLf & DateValue(DateClicked)

End Sub

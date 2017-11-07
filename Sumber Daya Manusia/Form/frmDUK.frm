VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDUK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Urut Kepangkatan Pegawai"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13830
   Icon            =   "frmDUK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   13830
   Begin MSFlexGridLib.MSFlexGrid fgLooping 
      Height          =   5775
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   10186
      _Version        =   393216
      Appearance      =   0
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
      Height          =   450
      Left            =   12120
      TabIndex        =   8
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Ceta&k"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10440
      TabIndex        =   7
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   13575
      Begin VB.OptionButton optsemua 
         Caption         =   "Semua"
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
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optRendah 
         Caption         =   "Terendah"
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
         Left            =   2400
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optTinggi 
         Caption         =   "Tertinggi"
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
         Left            =   1320
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Kriteria Urut"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         TabIndex        =   3
         Top             =   240
         Width           =   7935
         Begin VB.Frame Frame3 
            Caption         =   "Tanggal Lahir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4920
            TabIndex        =   17
            Top             =   120
            Width           =   2895
            Begin VB.OptionButton optThn 
               Caption         =   "Tahun"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1920
               TabIndex        =   20
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optBln 
               Caption         =   "Bulan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1080
               TabIndex        =   19
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optTgl 
               Caption         =   "Tanggal"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.OptionButton optJabatan 
            Caption         =   "Jabatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optUsia 
            Caption         =   "Batas Usia Pensiun"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optPangkat 
            Caption         =   "Pangkat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optPendidikan 
            Caption         =   "Pendidikan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   0
         Top             =   510
         Width           =   3495
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3960
         TabIndex        =   1
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Parameter"
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1680
      End
   End
   Begin MSDataGridLib.DataGrid dgPegawai 
      Height          =   5775
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
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
   Begin VB.Label lblJumData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data 0/0"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   8160
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDUK.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12000
      Picture         =   "frmDUK.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDUK.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDUK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilter As String
Dim rsb As New ADODB.recordset

Private Sub subLoadDataPegawai()
    On Error GoTo hell
    'Data Error Karena link Pada View tidak dapat menemukan link ke tabel pendidikan
    'buat view sementara
    
'    strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, dbo.V_DUKnew.Pangkat, dbo.V_DUKnew.TMTP, dbo.V_DUKnew.Golongan, dbo.V_DUKnew.TMTG, dbo.V_DUKnew.Jabatan, dbo.V_DUKnew.TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'    "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, dbo.V_DUKnew.NoUrutPangkat, dbo.V_DUKnew.NoUrutGolongan, " & _
'    "dbo.V_DUKnew.NoUrutJabatan , dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where NoUrutAlamat is NULL or NoUrutAlamat='001' " & strFilter



'//yayang.agus 2014-08-13
'    strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
    "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
    "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
    
    strSQL = "select * from v_duknew2_1 " & strFilter
    
'//


    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set dgPegawai.DataSource = rsb
    lblJumData.Caption = "Data " & dgPegawai.Bookmark & "/" & dgPegawai.ApproxCount
'//yayang.agus 2014-08-13
    With dgPegawai
        '.Columns("IdPegawai").Width = 0
        .Columns("Nama").Width = 3000
        .Columns("NIP").Width = 1000
        '.Columns("Pangkat").Width = 0
        .Columns("TMTP").Width = 1100
        .Columns("Golongan").Width = 1000
        .Columns("TMTG").Width = 1100
        .Columns("Jabatan").Width = 3000
        .Columns("TMTJ").Width = 1100
        '.Columns("Usia").Width = 0
        .Columns("Tanggal").Width = 0
        .Columns("Bulan").Width = 0
        .Columns("Tahun").Width = 0
        .Columns("NoUrutPangkat").Width = 0
        .Columns("NoUrutJabatan").Width = 0
        .Columns("NoUrutGolongan").Width = 0
        '.Columns("TglLahir").Width = 0
        .Columns("TempatLahir").Width = 1500
        .Columns("Alamat Lengkap").Width = 3000
        .Columns("TglLulus").Width = 1100
        .Columns("KdPendidikan").Width = 0
        .Columns("NoUrutAlamat").Width = 0
        .Columns("NoUrutPendidikan").Width = 0
        '.Columns("Pendidikan").Width = 0
        
        
'        .Columns("NoUrutPangkat").Width = 0
'        .Columns("NoUrutPendidikan").Width = 0
'        .Columns("NoUrutJabatan").Width = 0
'        .Columns("NoUrutGolongan").Width = 0
'        .Columns("Tanggal").Width = 0
'        .Columns("Bulan").Width = 0
'        .Columns("Tahun").Width = 0
'        .Columns("Usia").Width = 500
'        .Columns("TglLahir").Width = 1100
'        .Columns("Pendidikan").Width = 1200
'        .Columns("TMTJ").Width = 1100
'        .Columns("TMTP").Width = 1100
'        .Columns("TMTG").Width = 1100
'        .Columns("Jabatan").Width = 3000
'        .Columns("Golongan").Width = 1200
'        .Columns("Pangkat").Width = 2000
'        .Columns("NIP").Width = 1100
'        .Columns("Nama").Width = 3000
'        .Columns("Alamat Lengkap").Width = 3000
    End With
'//
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadPegawai()
On Error GoTo hell
'    strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, dbo.V_DUKnew.Pangkat, dbo.V_DUKnew.TMTP, dbo.V_DUKnew.Golongan, dbo.V_DUKnew.TMTG, dbo.V_DUKnew.Jabatan, dbo.V_DUKnew.TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'    "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, dbo.V_DUKnew.NoUrutPangkat, dbo.V_DUKnew.NoUrutGolongan, " & _
'    "dbo.V_DUKnew.NoUrutJabatan , dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan "

'    strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'    "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'    "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
    
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set dgPegawai.DataSource = rsb
    lblJumData.Caption = "Data " & dgPegawai.Bookmark & "/" & dgPegawai.ApproxCount

    With dgPegawai
        .Columns("NoUrutPangkat").Width = 0
        .Columns("NoUrutPendidikan").Width = 0
        .Columns("NoUrutJabatan").Width = 0
        .Columns("NoUrutGolongan").Width = 0
        .Columns("Tanggal").Width = 0
        .Columns("Bulan").Width = 0
        .Columns("Tahun").Width = 0
        .Columns("Usia").Width = 500
        .Columns("TglLahir").Width = 1100
        .Columns("Pendidikan").Width = 1200
        .Columns("TMTJ").Width = 1100
        .Columns("TMTP").Width = 1100
        .Columns("TMTG").Width = 1100
        .Columns("Jabatan").Width = 3000
        .Columns("Golongan").Width = 1200
        .Columns("Pangkat").Width = 2000
        .Columns("NIP").Width = 1100
        .Columns("Nama").Width = 3000
        .Columns("Alamat Lengkap").Width = 3000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Public Sub cmdCari_Click() '//yayang.agus 2014-08-13
    On Error GoTo errLoad
    
'    If txtParameter.Text = "" Then
'        'Call subLoadPegawai
'        Exit Sub
'    End If
    
    If optUsia.Value = True Then
        If optTinggi.Value = True Then
            strFilter = " WHERE Usia like '%" & txtParameter.Text & "%' order by Tahun ASC"
            
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Usia like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
            
        ElseIf optRendah.Value = True Then
            strFilter = " WHERE Usia like '%" & txtParameter.Text & "%' order by Tahun Desc"
            
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Usia like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
            
        ElseIf optSemua.Value = True Then
            strFilter = " WHERE Nama like '%" & txtParameter.Text & "%' order by Nama ASC"
            
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' AND Nama like '%" & txtParameter.Text & "%' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
            
        End If
    ElseIf optPangkat.Value = True Then
        If optTinggi.Value = True Then
            strFilter = " WHERE Pangkat like '%" & txtParameter.Text & "%' ORDER BY NoUrutPangkat ASC"

'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Pangkat like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
            
        ElseIf optRendah.Value = True Then
            strFilter = " WHERE Pangkat like '%" & txtParameter.Text & "%' ORDER BY NoUrutPangkat Desc"
            
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Pangkat like '%" & txtParameter.Text & "%' And NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
            
        ElseIf optSemua.Value = True Then
            strFilter = " WHERE Nama like '%" & txtParameter.Text & "%' order by Nama ASC"
        
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Nama like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
            
        End If
    ElseIf optJabatan.Value = True Then
        If optTinggi.Value = True Then
            strFilter = " WHERE Jabatan like '%" & txtParameter.Text & "%' ORDER BY NoUrutJabatan asc"
            
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Jabatan like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
            
        ElseIf optRendah.Value = True Then
            strFilter = " WHERE Jabatan like '%" & txtParameter.Text & "%' ORDER BY NoUrutJabatan desc"
        
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Jabatan like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
        
        ElseIf optSemua.Value = True Then
            strFilter = " WHERE Nama like '%" & txtParameter.Text & "%' order by Nama ASC"
            
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Nama like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
            
        End If
    ElseIf optTgl.Value = True Then
        strFilter = " WHERE Tanggal = '" & txtParameter.Text & "'"
        
'        strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'        "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'        "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Tanggal like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'        Call subLoadPegawai
        
    ElseIf optBln.Value = True Then
        strFilter = " WHERE Bulan = '" & txtParameter.Text & "'"
        
'        strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'        "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'        "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Bulan like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'        Call subLoadPegawai
'
    ElseIf optThn.Value = True Then
        strFilter = " WHERE Tahun = '" & txtParameter.Text & "'"
        
'        strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'        "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'        "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Tahun like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'        Call subLoadPegawai
        
    ElseIf optPendidikan.Value = True Then
        If optTinggi.Value = True Then
            strFilter = " WHERE Pendidikan like '%" & txtParameter.Text & "%' ORDER BY NoUrutPendidikan ASC"
            
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Pendidikan like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
            
        ElseIf optRendah.Value = True Then
            strFilter = " WHERE Pendidikan like '%" & txtParameter.Text & "%' ORDER BY NoUrutPendidikan Desc"
        
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Pendidikan like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
        
        ElseIf optSemua.Value = True Then
            strFilter = " WHERE Nama like '%" & txtParameter.Text & "%' order by Nama ASC"
        
'            strSQL = "SELECT dbo.V_DUKnew.Nama, dbo.V_DUKnew.NIP, MAX(dbo.V_DUKnew.Pangkat) AS Pangkat, MAX(dbo.V_DUKnew.TMTP) AS TMTP, MAX(dbo.V_DUKnew.Golongan) AS Golongan, MAX(dbo.V_DUKnew.TMTG) AS TMTG, max(dbo.V_DUKnew.Jabatan) AS Jabatan, MAX(dbo.V_DUKnew.TMTJ) AS TMTJ, dbo.Pendidikan.Pendidikan, dbo.V_DUKnew.TglLahir, dbo.V_DUKnew.Tanggal, dbo.V_DUKnew.Bulan, dbo.V_DUKnew.Tahun, " & _
'            "dbo.V_DUKnew.Usia, dbo.Pendidikan.NoUrut AS NoUrutPendidikan, MAX(dbo.V_DUKnew.NoUrutPangkat) AS NoUrutPangkat, MAX(dbo.V_DUKnew.NoUrutGolongan) AS NoUrutGolongan, " & _
'            "MAX(dbo.V_DUKnew.NoUrutJabatan) AS NoUrutJabatan, dbo.V_DUKnew.[Alamat Lengkap] FROM dbo.V_DUKnew LEFT OUTER JOIN dbo.Pendidikan ON dbo.V_DUKnew.KdPendidikan = dbo.Pendidikan.KdPendidikan where Nama like '%" & txtParameter.Text & "%' AND NoUrutJabatan is NULL AND NoUrutAlamat is NULL or NoUrutAlamat='001' GROUP BY Nama,NIP,TglLahir,Tanggal,Bulan,Tahun,Usia,[Alamat Lengkap],Pendidikan,NoUrut " & strFilter
'            Call subLoadPegawai
            
        End If
    End If

    Call subLoadDataPegawai
    If rsb.RecordCount = 0 Then Exit Sub
    Exit Sub
errLoad:
    Call msubPesanError
End Sub '//

Private Sub cmdCetak_Click()
    On Error GoTo hell
    Dim pesan As VbMsgBoxResult
'    cmdCetak.Enabled = False
'    If rsb.RecordCount = 0 Then
'        MsgBox "Tidak Ada Data", vbExclamation, "Validasi"
'        cmdCetak.Enabled = True
'        Exit Sub
'    End If

    '"tempDUK"
    'SELECT     no, Nama, Nip, Pangkat1, Pangkat2, Jabatan1, Jabatan2, MasaKerja1, MasaKerja2, LatihanJabatan1,
    '            LatihanJabatan2, LatihanJabatan3, Pendidikan1, Pendidikan2,
    '           Pendidikan3 , Usia, MutasiKerja
    'From tempDUK

    Dim no, Nama, Nip, Pangkat1, Pangkat2, Jabatan1, Jabatan2, MasaKerja1, MasaKerja2, LatihanJabatan1, LatihanJabatan2, LatihanJabatan3, Pendidikan1, Pendidikan2, Pendidikan3, Usia, MutasiKerja As String
    Dim ii As Integer
    Dim BRS As Integer
    Dim BRS_AWAL As Integer
    
    strsqlx = "delete from tempDUK"
    Call msubRecFO(rsx, strsqlx)
    strSQL = "select * from v_duknew2 " & strFilter
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        fgLooping.Rows = 1
        fgLooping.Cols = 20
        BRS = 0
        BRS_AWAL = 1
        For ii = 0 To rs.RecordCount - 1
            fgLooping.Rows = fgLooping.Rows + 1
'            BRS = fgLooping.Rows - 1
            no = no + 1
            Nama = IIf(IsNull(rs!Nama), "", rs!Nama) 'rs!Nama
            Nip = IIf(IsNull(rs!Nip), "", rs!Nip) 'rs!Nip
            
            'BRS = BRS + 1
            BRS_AWAL = fgLooping.Rows - 1
            
            'For i = 1 To fgLooping.Rows - 1
                fgLooping.TextMatrix(BRS_AWAL, 1) = no
                fgLooping.TextMatrix(BRS_AWAL, 2) = Nama
                fgLooping.TextMatrix(BRS_AWAL, 3) = Nip
            'Next
            
            strsqlxx = "SELECT NamaGolongan,datepart(year,TglSK) as Thn  FROM v_RiwayatPangkat WHERE     IdPegawai = '" & rs(0) & "'"
            Call msubRecFO(rsxx, strsqlxx)
            If rsxx.RecordCount <> 0 Then
                BRS = BRS_AWAL - 1
                For i = 1 To rsxx.RecordCount
                    BRS = BRS + 1
                    If fgLooping.Rows - 1 < BRS Then fgLooping.Rows = BRS + 1
                    fgLooping.TextMatrix(BRS, 4) = IIf(IsNull(rsxx(0)), "", rsxx(0)) 'rsxx(0)
                    fgLooping.TextMatrix(BRS, 5) = IIf(IsNull(rsxx(1)), "", rsxx(1)) 'rsxx(1)
                    rsxx.MoveNext
                Next
            End If
            DoEvents
            strsqlxx = "select namajabatan,datepart(year,tglsk) as Thn from v_RiwayatJabatan where IdPegawai = '" & rs(0) & "'"
            Call msubRecFO(rsxx, strsqlxx)
            If rsxx.RecordCount <> 0 Then
                BRS = BRS_AWAL - 1
                For i = 0 To rsxx.RecordCount - 1
                    BRS = BRS + 1
                    If fgLooping.Rows - 1 < BRS Then fgLooping.Rows = BRS + 1
                    fgLooping.TextMatrix(BRS, 6) = IIf(IsNull(rsxx(0)), "", rsxx(0)) 'rsxx(0)
                    fgLooping.TextMatrix(BRS, 7) = IIf(IsNull(rsxx(1)), "", rsxx(1)) 'rsxx(1)
                    rsxx.MoveNext
                Next
            End If
            
            'For i = 1 To fgLooping.Rows - 1
                fgLooping.TextMatrix(BRS_AWAL, 8) = Val(IIf(IsNull(rs!MasaKerjaBln), "", rs!MasaKerjaBln)) \ 12
                fgLooping.TextMatrix(BRS_AWAL, 9) = Val(IIf(IsNull(rs!MasaKerjaBln), "", rs!MasaKerjaBln)) Mod 12
            'Next
            
            strsqlxx = "select NamaDiklat,datepart(year,TglMulai) as Thn,jmlJam  from V_RiwayatDiklat where IdPegawai = '" & rs(0) & "'"
            Call msubRecFO(rsxx, strsqlxx)
            If rsxx.RecordCount <> 0 Then
                BRS = BRS_AWAL - 1
                For i = 0 To rsxx.RecordCount - 1
                    BRS = BRS + 1
                    If fgLooping.Rows - 1 < BRS Then fgLooping.Rows = BRS + 1
                    fgLooping.TextMatrix(BRS, 10) = IIf(IsNull(rsxx(0)), "", rsxx(0)) 'rsxx(0)
                    fgLooping.TextMatrix(BRS, 11) = IIf(IsNull(rsxx(1)), "", rsxx(1)) '
                    fgLooping.TextMatrix(BRS, 12) = IIf(IsNull(rsxx(2)), "", rsxx(2)) '
                    rsxx.MoveNext
                Next
            End If
            
            DoEvents
            strsqlxx = "select namapendidikan,datepart(year,tgllulus) as Thn,pendidikan from v_RiwayatPendidikan where IdPegawai = '" & rs(0) & "'"
            Call msubRecFO(rsxx, strsqlxx)
            If rsxx.RecordCount <> 0 Then
                BRS = BRS_AWAL - 1
                For i = 0 To rsxx.RecordCount - 1
                    BRS = BRS + 1
                    If fgLooping.Rows - 1 < BRS Then fgLooping.Rows = BRS + 1
                    fgLooping.TextMatrix(BRS, 13) = IIf(IsNull(rsxx(0)), "", rsxx(0)) 'rsxx(0)
                    fgLooping.TextMatrix(BRS, 14) = IIf(IsNull(rsxx(1)), "", rsxx(1)) '
                    fgLooping.TextMatrix(BRS, 15) = IIf(IsNull(rsxx(2)), "", rsxx(2)) '
                    rsxx.MoveNext
                Next
            End If
            
            'For i = 1 To fgLooping.Rows - 1
                fgLooping.TextMatrix(BRS_AWAL, 16) = IIf(IsNull(rs!Usia), "", rs!Usia) 'rsxx(0)
            'Next
            
            DoEvents
            'strsqlxx = "select Jabatan,Tempat,Tahun  from RiwayatMutasiPegawai where IdPegawai = '" & rs(0) & "'"
            strsqlxx = "SELECT JabatanPosisi, NamaPerusahaan,  datepart(year,TglMulai) FROM RiwayatPekerjaan where IdPegawai = '" & rs(0) & "'"
            Call msubRecFO(rsxx, strsqlxx)
            If rsxx.RecordCount <> 0 Then
                BRS = BRS_AWAL - 1
                For i = 0 To rsxx.RecordCount - 1
                    BRS = BRS + 1
                    If fgLooping.Rows - 1 < BRS Then fgLooping.Rows = BRS + 1
                    fgLooping.TextMatrix(BRS, 17) = IIf(IsNull(rsxx(0)), "", rsxx(0)) & " di " & IIf(IsNull(rsxx(1)), "", rsxx(1)) & " Tahun " & IIf(IsNull(rsxx(2)), "", rsxx(2))
                    rsxx.MoveNext
                Next
            End If
            
            'Pangkat1 = ""
            'Pangkat2 = ""
            'Jabatan1 = ""
'            'Jabatan2 = ""
'            MasaKerja1 = rs!MasaKerjaThn
'            MasaKerja2 = rs!MasaKerjaBln
'            LatihanJabatan1 = ""
'            LatihanJabatan2 = ""
'            LatihanJabatan3 = ""
'            Pendidikan1 = ""
'            Pendidikan2 = ""
'            Pendidikan3 = ""
''            Usia = rs!Usia
'            MutasiKerja = ""

            
            lblJumData.Caption = ii & "/" & rs.RecordCount - 1
            rs.MoveNext
        Next
    End If
    
    Dim brsNama As String
    Dim namaa As String
    
    For i = 1 To fgLooping.Rows - 1
        If fgLooping.TextMatrix(i, 2) <> "" Then
            brsNama = i
            namaa = fgLooping.TextMatrix(i, 2)
        End If
        fgLooping.TextMatrix(i, 2) = namaa
    Next
    
    With fgLooping
        For i = 1 To .Rows - 1
            strsqlx = "insert into tempDUK values (" & _
                    "'" & .TextMatrix(i, 1) & "','" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 3) & "','" & .TextMatrix(i, 4) & "','" & .TextMatrix(i, 5) & "','" & .TextMatrix(i, 6) & "','" & .TextMatrix(i, 7) & "'," & _
                    "'" & .TextMatrix(i, 8) & "','" & .TextMatrix(i, 9) & "','" & .TextMatrix(i, 10) & "','" & .TextMatrix(i, 11) & "','" & .TextMatrix(i, 12) & "'," & _
                    "'" & .TextMatrix(i, 13) & "','" & .TextMatrix(i, 14) & "','" & .TextMatrix(i, 15) & "','" & .TextMatrix(i, 16) & "','" & .TextMatrix(i, 17) & "'" & _
                    ")"
            Call msubRecFO(rsx, strsqlx)
        Next
    End With
    
    strSQL = "select * from tempDUK"
    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"
    FrmCetakDUKPegawai.Show
    cmdCetak.Enabled = True
hell:
'Resume 0
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgPegawai_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPegawai
    WheelHook.WheelHook dgPegawai
End Sub

Private Sub dgPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    lblJumData.Caption = "Data " & dgPegawai.Bookmark & "/" & dgPegawai.ApproxCount
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    optTinggi.Enabled = True
    optRendah.Enabled = True
    optSemua.Enabled = True
    optUsia.Value = True
    optSemua.Value = True
    strFilter = "ORDER by Nama ASC"
    Call subLoadDataPegawai

End Sub

Private Sub optBln_Click()
    optUsia.Value = False
    optPangkat.Value = False
    optJabatan.Value = False
    optPendidikan.Value = False
    optTinggi.Enabled = False
    optRendah.Enabled = False
    optSemua.Enabled = False
    
    optUsia.Enabled = True
    optPangkat.Enabled = True
    optJabatan.Enabled = True
    optPendidikan.Enabled = True
    
    Label1.Caption = "Masukkan Bulan Lahir Format (MM)"

End Sub

Private Sub optJabatan_Click()
    optTgl.Value = False
    optBln.Value = False
    optThn.Value = False
    optTinggi.Enabled = True
    optRendah.Enabled = True
    optSemua.Enabled = True
    Label1.Caption = "Masukkan Nama Jabatan"
End Sub

Private Sub optPangkat_Click()
    optTgl.Value = False
    optBln.Value = False
    optThn.Value = False
    optTinggi.Enabled = True
    optRendah.Enabled = True
    optSemua.Enabled = True
    Label1.Caption = "Masukkan Nama Pangkat"
End Sub

Private Sub optPendidikan_Click()
    optTgl.Value = False
    optBln.Value = False
    optThn.Value = False
    optTinggi.Enabled = True
    optRendah.Enabled = True
    optSemua.Enabled = True
    Label1.Caption = "Masukkan Pendidikan"
End Sub

Private Sub optRendah_Click()
    optUsia.Enabled = True
    optPangkat.Enabled = True
    optJabatan.Enabled = True
    optPendidikan.Enabled = True
    Label1.Caption = "Masukan nama sesuai Kriteria Urut"
End Sub

Private Sub optsemua_Click()
    optPendidikan.Value = True
    optUsia.Enabled = False
    optPangkat.Enabled = False
    optJabatan.Enabled = False
    optPendidikan.Enabled = False
     If optSemua.Value = True Then Label1.Caption = "Masukkan Nama Pegawai"
End Sub

Private Sub optTgl_Click()
    optUsia.Value = False
    optPangkat.Value = False
    optJabatan.Value = False
    optPendidikan.Value = False
    optTinggi.Enabled = False
    optRendah.Enabled = False
    optSemua.Enabled = False
    
    optUsia.Enabled = True
    optPangkat.Enabled = True
    optJabatan.Enabled = True
    optPendidikan.Enabled = True
    
    Label1.Caption = "Masukkan Tanggal Lahir Format (DD)"

End Sub

Private Sub optThn_Click()
    optUsia.Value = False
    optPangkat.Value = False
    optJabatan.Value = False
    optPendidikan.Value = False
    optTinggi.Enabled = False
    optRendah.Enabled = False
    optSemua.Enabled = False
    
    optUsia.Enabled = True
    optPangkat.Enabled = True
    optJabatan.Enabled = True
    optPendidikan.Enabled = True
    
    Label1.Caption = "Masukkan Tahun Lahir  Format (YY)"

End Sub

Private Sub optTinggi_Click()
    optUsia.Enabled = True
    optPangkat.Enabled = True
    optJabatan.Enabled = True
    optPendidikan.Enabled = True
    Label1.Caption = "Masukan nama sesuai Kriteria Urut"
End Sub

Private Sub optUsia_Click()
    optTgl.Value = False
    optBln.Value = False
    optThn.Value = False
    optTinggi.Enabled = True
    optRendah.Enabled = True
    optSemua.Enabled = True
    Label1.Caption = "Masukkan Usia"
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
        If optUsia.Value = True Then
            Call cmdCari_Click
            Call SetKeyPressToNumber(KeyAscii)
        End If
        If optPangkat.Value = True Then
            Call cmdCari_Click
            Call SetKeyPressToChar(KeyAscii)
        End If
        If optJabatan.Value = True Then
            Call cmdCari_Click
            Call SetKeyPressToChar(KeyAscii)
        End If
        If optPendidikan.Value = True Then
            Call cmdCari_Click
            'Call SetKeyPressToChar(KeyAscii)
        End If
        
        If optTgl.Value = True Then
            Call cmdCari_Click
            Call SetKeyPressToNumber(KeyAscii)
        End If
        If optBln.Value = True Then
            Call cmdCari_Click
            Call SetKeyPressToNumber(KeyAscii)
        End If
        If optThn.Value = True Then
            Call cmdCari_Click
            Call SetKeyPressToNumber(KeyAscii)
        End If
        If optUsia.Value = True Then
            Call cmdCari_Click
            Call SetKeyPressToNumber(KeyAscii)
        End If
    Call subLoadDataPegawai
End Sub

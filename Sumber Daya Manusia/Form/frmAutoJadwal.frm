VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAutoJadwal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jadwal Automatis"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   Icon            =   "frmAutoJadwal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10335
   Begin VB.CommandButton cmdInsertRow 
      Caption         =   "&InsertRow"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCol 
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtRow 
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdBuatJadwal 
      Caption         =   "&Buat Jadwal"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame frIsiList 
      Height          =   2175
      Left            =   6600
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3255
         Begin VB.ComboBox cmbSebanyak 
            Height          =   315
            ItemData        =   "frmAutoJadwal.frx":0CCA
            Left            =   1320
            List            =   "frmAutoJadwal.frx":0CD4
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   855
         End
         Begin MSDataListLib.DataCombo dcShiftKerja 
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   240
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "hari"
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
            Left            =   2280
            TabIndex        =   19
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Sebanyak"
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
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Shift Terakhir"
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
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdLanjut 
         Caption         =   "&Lanjutkan"
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   10215
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
      Begin MSDataListLib.DataCombo dcTempatBertugas 
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcShiftKerja2 
         Height          =   315
         Left            =   3840
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
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
      Begin VB.CheckBox chksmua 
         Caption         =   "Pilih semua karyawan"
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
         Left            =   3600
         TabIndex        =   24
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   390
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1920
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgJadwal 
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6165
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   8640
      TabIndex        =   15
      Top             =   7920
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvjadwalkerja 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Pegawai"
         Object.Width           =   13229
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Shift Terakhir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sebanyak (hari)"
         Object.Width           =   2540
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
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
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "frmAutoJadwal.frx":0CDE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   945
      Left            =   8760
      Picture         =   "frmAutoJadwal.frx":369F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmAutoJadwal.frx":4427
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmAutoJadwal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intJumHari As Integer
Private arrTemShiftKerja(35) As String
Dim strHari As String

Private Sub subSetListPegawai()
    Dim intLebarCol As Integer

    intLebarCol = Me.lvjadwalkerja.Width / 3
    Me.lvjadwalkerja.ColumnHeaders.Item(1).Width = intLebarCol + 1500
    Me.lvjadwalkerja.ColumnHeaders.Item(2).Width = intLebarCol - 300 - 500
    Me.lvjadwalkerja.ColumnHeaders.Item(3).Width = intLebarCol - 1000
End Sub

Private Sub subSetDC()
    strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan order by NamaRuangan"
    Call msubDcSource(dcTempatBertugas, rs, strSQL)

    strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja_New order by KdShift"
    Call msubDcSource(dcShiftKerja, rs, strSQL)

    strSQL = "Select IdShift, Dinaskerja From DinasKerja"
    Call msubDcSource(dcShiftKerja2, rs, strSQL)
End Sub

Private Sub subLoadNamaPegawai()
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
End Sub

Private Sub subSetFGJadwal()
    Dim i As Integer
    Dim itm As ListItem

    Call totalHari(Me.DTPicker1.Month, Me.DTPicker1.Year)
    intJumHari = TtlHari

    With Me.fgJadwal
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
    End With
End Sub

Private Sub subBuatJadwal()
    Dim itm As ListItem
    Dim intShiftTerakhir As Integer, intShiftSekarang As Integer
    Dim strShiftTerakhir As String, strShiftSekarang As String
    Dim intBanyakShiftTerakhir As Integer
    Dim intBanyakShiftSekarang As Integer
    Dim intRowShiftKerja As Integer
    Dim intRowNama As Integer
    Dim intRowInsert As Integer, intRowCetak As Integer
    Dim i As Integer, j As Integer, k As Integer, l As Integer

    Call subSetFGJadwal

    For Each itm In Me.lvjadwalkerja.ListItems

        If itm.Checked Then

            If itm.SubItems(1) = "" Or itm.SubItems(2) = "" Then
                MsgBox "Data kurang!", vbCritical, "Error"
                itm.Selected = True
                Me.lvjadwalkerja.SetFocus
                Exit Sub
            End If
            strShiftTerakhir = itm.SubItems(1)
            intShiftTerakhir = funcConvertShiftKerjaToNumber(strShiftTerakhir)

            intBanyakShiftTerakhir = CInt(itm.SubItems(2))
            If intBanyakShiftTerakhir = 2 Then
                intBanyakShiftSekarang = 2
                intShiftSekarang = intShiftTerakhir + 1
            Else
                intBanyakShiftSekarang = 1
                intShiftSekarang = intShiftTerakhir
            End If

            With Me.fgJadwal
                For j = 1 To intJumHari
                    If intShiftSekarang > 4 Then intShiftSekarang = 1
                    If intShiftSekarang = 4 Then
                        GoTo jump
                    End If
                    strShiftSekarang = funcConvertNumberToShiftKerja(intShiftSekarang)
                    intRowShiftKerja = funcCariRowShiftKerja(strShiftSekarang)
                    intRowNama = funcCariRowShiftKerja(itm.Text, True, intRowShiftKerja)
                    If intRowNama > 0 Then
                        If .TextMatrix(intRowNama, 0) = strShiftSekarang Then
                            intRowShiftKerja = intRowNama
                        End If
                    End If
                    If .TextMatrix(intRowShiftKerja, 1) = "" Then
                        .TextMatrix(intRowShiftKerja, 1) = itm.Text
                        intRowCetak = intRowShiftKerja
                    ElseIf .TextMatrix(intRowShiftKerja, 1) = itm.Text Then
                        intRowCetak = intRowShiftKerja
                    Else
                        intRowInsert = intRowShiftKerja + 1
                        .AddItem strShiftSekarang, intRowInsert
                        .TextMatrix(intRowInsert, 1) = itm.Text
                        intRowCetak = intRowInsert
                    End If

                    .TextMatrix(intRowCetak, j + 1) = "X"
                    .row = intRowCetak
                    .Col = j + 1
                    .CellAlignment = flexAlignCenterCenter
                    .CellFontSize = 10
                    .CellFontBold = True
                    .CellForeColor = vbRed

jump:
                    intBanyakShiftSekarang = intBanyakShiftSekarang - 1
                    If intBanyakShiftSekarang = 0 Then
                        intBanyakShiftSekarang = 2
                        intShiftSekarang = intShiftSekarang + 1
                    End If
                Next
            End With
        End If
    Next
End Sub

Private Sub subBuatJadwalBaru()
    Dim itm As ListItem
    Dim intShiftTerakhir As Integer, intShiftSekarang As Integer
    Dim strShiftTerakhir As String, strShiftSekarang As String
    Dim intBanyakShiftTerakhir As Integer
    Dim intBanyakShiftSekarang As Integer
    Dim i As Integer
    Dim strKodeShift As String

    Call subSetFGJadwal

    For Each itm In Me.lvjadwalkerja.ListItems
        If itm.Checked Then
            With Me.fgJadwal

                strsqlx = "SELECT ConvertIdPegawaiToShift.IdPegawai,DataPegawai.NamaLengkap,ConvertIdPegawaiToShift.IdShift From ConvertIdPegawaiToShift INNER JOIN DataPegawai ON ConvertIdPegawaiToShift.IdPegawai = DataPegawai.IdPegawai WHERE (NamaLengkap = '" & itm.Text & "')"
                Set rsx = Nothing
                Call msubRecFO(rsx, strsqlx)

                If rsx.EOF = True Then
                    MsgBox "Ada pegawai yang belum ditentukan shift nya!", vbCritical, "Error"
                    Exit Sub
                End If
                If rsx(2).Value <> "02" Then

                    If itm.SubItems(1) = "" Or itm.SubItems(2) = "" Then
                        MsgBox "Data kurang!", vbCritical, "Error"
                        itm.Selected = True
                        Me.lvjadwalkerja.SetFocus
                        Exit Sub
                    End If

                    .row = .Rows - 1
                    .TextMatrix(.row, 0) = itm.Text

                    strShiftTerakhir = itm.SubItems(1)
                    intShiftTerakhir = funcConvertShiftKerjaToNumber(strShiftTerakhir)

                    intBanyakShiftTerakhir = CInt(itm.SubItems(2))
                    If intBanyakShiftTerakhir = 2 Then
                        intBanyakShiftSekarang = 2
                        intShiftSekarang = intShiftTerakhir + 1
                    Else
                        intBanyakShiftSekarang = 1
                        intShiftSekarang = intShiftTerakhir
                    End If

                    For i = 1 To intJumHari 'isi absensi
                        If intShiftSekarang > 4 Then intShiftSekarang = 1

                        strShiftSekarang = funcConvertNumberToShiftKerja(intShiftSekarang)
                        strKodeShift = UCase$(Left(strShiftSekarang, 1))

                        .Col = i
                        .TextMatrix(.row, i) = strKodeShift
                        .CellAlignment = flexAlignCenterCenter
                        .CellFontBold = True

                        intBanyakShiftSekarang = intBanyakShiftSekarang - 1
                        If intBanyakShiftSekarang = 0 Then
                            intBanyakShiftSekarang = 2
                            intShiftSekarang = intShiftSekarang + 1
                        End If
                    Next
                    .Rows = .Rows + 1

                Else

                    If itm.SubItems(1) = "" Then
                        MsgBox "Data kurang!", vbCritical, "Error"
                        itm.Selected = True
                        Me.lvjadwalkerja.SetFocus
                        Exit Sub
                    End If

                    .row = .Rows - 1
                    .TextMatrix(.row, 0) = itm.Text

                    strShiftTerakhir = itm.SubItems(1)
                    intShiftTerakhir = funcConvertShiftKerjaToNumber(strShiftTerakhir)

                    intBanyakShiftTerakhir = 1
                    If intBanyakShiftTerakhir = 2 Then
                        intBanyakShiftSekarang = 2
                        intShiftSekarang = intShiftTerakhir + 1
                    Else
                        intBanyakShiftSekarang = 1
                        intShiftSekarang = intShiftTerakhir
                    End If

                    For i = 1 To intJumHari 'isi absensi
                        If intShiftSekarang > 4 Then intShiftSekarang = 1

                        strShiftSekarang = funcConvertNumberToShiftKerja(intShiftSekarang)
                        strKodeShift = UCase$(Left(strShiftSekarang, 1))

                        .Col = i
                        .TextMatrix(.row, i) = "P"
                        .CellAlignment = flexAlignCenterCenter
                        .CellFontBold = True

                        Call totalHari(Me.DTPicker1.Month, Me.DTPicker1.Year)
                        intJumHari = TtlHari

                        Me.DTPicker1.Day = i
                        strHari = WeekdayName(Weekday(Me.DTPicker1.Value), , vbSunday)

                        .TextMatrix(0, i) = Format(Me.DTPicker1.Value, "MMMM yyyy")
                        .TextMatrix(1, i) = i

                        If strHari = "Sabtu" Or strHari = "Minggu" Then
                            .TextMatrix(.row, i) = "L"
                            .CellAlignment = flexAlignCenterCenter
                            .CellFontBold = True
                        End If

                        strSQL = "SELECT [Hari Libur] FROM v_tanggal" & _
                        " WHERE DAY(Tanggal)='" & i & "'" & _
                        " AND MONTH(Tanggal)='" & Month(Me.DTPicker1.Value) & "'" & _
                        " AND YEAR(Tanggal)='" & Year(Me.DTPicker1.Value) & "'"
                        Call msubRecFO(rs, strSQL)

                        If Not rs.EOF Then
                            If Not IsNull(rs.Fields.Item("Hari Libur").Value) Then
                                .TextMatrix(.row, i) = "L"
                                .CellAlignment = flexAlignCenterCenter
                                .CellFontBold = True
                            End If
                        End If

                        intBanyakShiftSekarang = intBanyakShiftSekarang - 1
                        If intBanyakShiftSekarang = 0 Then
                            intBanyakShiftSekarang = 2
                            intShiftSekarang = intShiftSekarang + 1
                        End If
                    Next
                    .Rows = .Rows + 1

                End If

            End With
        End If
    Next
End Sub

Private Function funcConvertShiftKerjaToNumber(ByVal ShiftKerja As String) As Integer
    strSQL = "SELECT KdShift FROM ShiftKerja_New WHERE NamaShift='" & ShiftKerja & "'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        funcConvertShiftKerjaToNumber = CInt(rs.Fields.Item("KdShift").Value)
    End If
End Function

Private Function funcConvertNumberToShiftKerja(ByVal NomorShift As Integer) As String
    Dim strNomorShift As String

    strNomorShift = "0" & CStr(NomorShift)
    strSQL = "SELECT NamaShift FROM ShiftKerja_New WHERE KdShift='" & strNomorShift & "' and namashift <> 'Reguler'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        funcConvertNumberToShiftKerja = rs.Fields.Item("NamaShift").Value
    '//yayang.agus 2014-08-22
    Else
        funcConvertNumberToShiftKerja = "LIBUR"
    '//
    End If
End Function

Private Function funcCariRowShiftKerja(ByVal strCari As String, Optional ByVal CariNama As Boolean, Optional ByVal CurIndex As Integer) As Integer
    Dim intMaxRow As Integer, i As Integer
    Dim intCol As Integer, intMulai As Integer

    With Me.fgJadwal
        If CariNama Then
            intCol = 1
            intMulai = CurIndex
        Else
            intCol = 0
            intMulai = 1
        End If
        intMaxRow = .Rows - 1
        For i = intMulai To intMaxRow
            If .TextMatrix(i, intCol) = strCari Then
                funcCariRowShiftKerja = i
                Exit For
            Else
                funcCariRowShiftKerja = 0
            End If
        Next
    End With
End Function

Private Sub subSimpanJadwal()
    On Error GoTo tangani

    Dim intMaxRow As Integer, intMaxCol As Integer
    Dim c As Integer, R As Integer
    Dim strKdTgl As String, strKdShift As String, strIDPegawai As String

    With Me.fgJadwal
        intMaxRow = .Rows - 1
        intMaxCol = .Cols - 1

        For c = 2 To intMaxCol
            Me.DTPicker1.Day = CInt(.TextMatrix(1, c))
            strSQL = "SELECT [Kode Tgl] FROM v_tanggal" & _
            " WHERE DAY(Tanggal)='" & Me.DTPicker1.Day & "'" & _
            " AND MONTH(Tanggal)='" & Me.DTPicker1.Month & "'" & _
            " AND YEAR(Tanggal)='" & Me.DTPicker1.Year & "'"
            Call msubRecFO(rs, strSQL)
            If Not rs.EOF Then
                strKdTgl = rs.Fields.Item("Kode Tgl").Value
            End If
            For R = 1 To intMaxRow
                If .TextMatrix(R, c) <> "" Then
                    strIDPegawai = funcCariIdPegawai(.TextMatrix(R, 1))
                    strKdShift = funcCariKodeShift(.TextMatrix(R, 0))
                    Call subJadwalKerja(strKdTgl, strIDPegawai, strKdShift)
                End If
            Next
        Next
    End With
    MsgBox "Simpan jadwal selesai.", vbInformation, "Perhatian"
    Exit Sub

tangani:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub subSimpanJadwalBaru()
    On Error GoTo tangani

    Dim intMaxRow As Integer, intMaxCol As Integer
    Dim c As Integer, R As Integer
    Dim strKdTgl As String, strKdShift As String, strIDPegawai As String

    With Me.fgJadwal
        intMaxRow = .Rows - 1
        intMaxCol = .Cols - 1

        For R = 2 To intMaxRow
            strIDPegawai = funcCariIdPegawai(.TextMatrix(R, 0))
            For c = 1 To intMaxCol
                Me.DTPicker1.Day = CInt(.TextMatrix(1, c))
                strSQL = "SELECT [Kode Tgl] FROM v_tanggal" & _
                " WHERE DAY(Tanggal)='" & Me.DTPicker1.Day & "'" & _
                " AND MONTH(Tanggal)='" & Me.DTPicker1.Month & "'" & _
                " AND YEAR(Tanggal)='" & Me.DTPicker1.Year & "'"
                Call msubRecFO(rs, strSQL)
                If Not rs.EOF Then
                    strKdTgl = rs.Fields.Item("Kode Tgl").Value
                End If
                strKdShift = funcCariKodeShift(.TextMatrix(R, c))
                Call subJadwalKerjaDetail(strIDPegawai, Me.DTPicker1.Day & "/" & Me.DTPicker1.Month & "/" & Me.DTPicker1.Year, strKdShift)
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

Public Sub subJadwalKerja(ByVal KdTgl As String, ByVal idpegawai As String, ByVal KdShift As String)
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdTgl", adVarChar, adParamInput, 3, KdTgl)
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
        .Parameters.Append .CreateParameter("JadwalKerja", adDate, adParamInput, , Format(tglJadwalKerja, "dd/MM/yyyy"))
        .Parameters.Append .CreateParameter("KdShift", adChar, adParamInput, 2, KdShift)

        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AU_JADWALKERJADETAIL"
        .Execute

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Sub

Private Sub subIsiFgJadwalKerja()
'//yayang.agus 2014-08-25
    Dim intRowNama As Integer, intColTgl As Integer
    Dim strNama As String, strTanggal As String, strShift As String
    Dim strKodeShift As String
    Dim i, j As Integer

    
'    For i = 2 To Me.fgJadwal.Rows - 1 ' - 2
    intRowNama = funcCariRowNama(rs.Fields.Item("Nama").Value)
        For j = 1 To Me.fgJadwal.Cols - 1
            Me.fgJadwal.TextMatrix(intRowNama, j) = "L"
        Next j
'    Next i
    While Not rs.EOF
        strNama = rs.Fields.Item("Nama").Value
        strShift = rs.Fields.Item("Shift").Value
        strTanggal = rs.Fields.Item("Tanggal").Value

        intRowNama = funcCariRowNama(strNama)
        intColTgl = funcCariColTanggal(Day(strTanggal))

        strKodeShift = UCase$(Left(strShift, 1))
        Me.fgJadwal.TextMatrix(intRowNama, intColTgl) = strKodeShift
        rs.MoveNext
    Wend
    '//
    
'    Dim intRowNama As Integer, intColTgl As Integer
'    Dim strNama As String, strTanggal As String, strShift As String
'    Dim strKodeShift As String
'    Dim i As Integer
'
'    strSQL = "SELECT * FROM V_JadwalKerjaNew" & _
'    " WHERE Ruangan='" & dcTempatBertugas.Text & "'" & _
'    " AND MONTH(Tanggal)='" & Me.DTPicker1.Month & "'" & _
'    " AND YEAR(Tanggal)='" & Me.DTPicker1.Year & "'" & _
'    " ORDER BY Tanggal"
'
'    Call msubRecFO(rs, strSQL)
'    While Not rs.EOF
'        strNama = rs.Fields.Item("Nama").Value
'        strShift = rs.Fields.Item("Shift").Value
'        strTanggal = rs.Fields.Item("Tanggal").Value
'
'        intRowNama = funcCariRowNama(strNama)
'        intColTgl = funcCariColTanggal(Day(strTanggal))
'
'        strKodeShift = UCase$(Left(strShift, 1))
'        Me.fgJadwal.TextMatrix(intRowNama, intColTgl) = strKodeShift
'        rs.MoveNext
'    Wend
End Sub

Private Function funcCariRowNama(ByVal NamaPegawai As String) As Integer
    Dim intMaxRow As Integer
    Dim i As Integer

    With Me.fgJadwal
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

    With Me.fgJadwal
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

Private Sub cmbSebanyak_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdLanjut.SetFocus
    End If
End Sub

Private Sub cmdBatal_Click()
    Me.frIsiList.Visible = False
    Me.lvjadwalkerja.SetFocus
End Sub

Private Sub cmdBuatJadwal_Click()
    Call subBuatJadwalBaru
End Sub

Private Sub cmdInsertRow_Click()
    With Me.fgJadwal
        .row = CInt(Me.txtRow.Text)
        .Col = CInt(Me.txtCol.Text)
        .AddItem "Coba", .row + 1
    End With
End Sub

Private Sub cmdLanjut_Click()
    If Me.dcShiftKerja.Text = "" Or Me.cmbSebanyak.Text = "" Then
        MsgBox "Data kurang!", vbCritical, "Validasi"
        Exit Sub
    End If
    With Me.lvjadwalkerja.SelectedItem
        .SubItems(1) = Me.dcShiftKerja.Text
        .SubItems(2) = Me.cmbSebanyak.Text
    End With
    Me.frIsiList.Visible = False
    Me.lvjadwalkerja.SetFocus
End Sub

Private Sub cmdSimpan_Click()
Dim i As Integer
    If dcTempatBertugas.Text <> "" Then
        If Periksa("datacombo", dcTempatBertugas, "Ruangan Tempat Bertugas Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcTempatBertugas.Text = "" Then
        If Periksa("datacombo", dcTempatBertugas, "Ruangan Tempat Bertugas Kosong") = False Then Exit Sub
    End If
    '//yayang.agus 2014-08-25
'    With Me.lvjadwalkerja.SelectedItem
'        If .SubItems(1) = "" Then
'            MsgBox "Nama Pegawai, shift terakhir dan hari harus diisi", vbInformation, "Validasi"
'            Exit Sub
'        End If
'    End With
    '//

    If fgJadwal.TextMatrix(2, i) = "" Then
        MsgBox "Tabel jadwal masih kosong", vbInformation, "Validasi"
        Exit Sub
    End If
    Call subSimpanJadwalBaru
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcShiftKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmbSebanyak.SetFocus
    End If
End Sub

Private Sub dcTempatBertugas_Change()
    Call subLoadNamaPegawai
    Call subSetFGJadwal
End Sub

Private Sub dcTempatBertugas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub DTPicker1_Change()
    Call subSetFGJadwal
    Call subLoadNamaPegawai
End Sub

Private Sub fgJadwal_Click()
    With Me.fgJadwal
        Me.txtCol.Text = .Col
        Me.txtRow.Text = .row
    End With
End Sub

Private Sub Form_Load()
    DTPicker1.Value = glDate
    Me.dcTempatBertugas.Text = frmJadwalKerja.dcTempatBertugas.Text
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    Call subSetListPegawai
    Call subSetDC
    Call subSetFGJadwal
End Sub


Private Sub lvjadwalkerja_DblClick()
    Call lvjadwalkerja_KeyPress(13)
End Sub

Private Sub lvjadwalkerja_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim strShiftTerakhir As String
    Dim strTempShift As String
    Dim intBanyakShiftTerakhir As Integer
    Dim i, j As Integer

    If Not Item.Checked Then
        Item.SubItems(1) = ""
        Item.SubItems(2) = ""
        
        j = funcCariRowNama(Item.Text)
        If j <> 0 Then
            If Me.fgJadwal.Rows = 3 Then
                For i = 0 To Me.fgJadwal.Cols - 1
                    Me.fgJadwal.TextMatrix(j, i) = ""
                Next i
            Else
                fgJadwal.RemoveItem (j)       '//yayang.agus 2014-08-25
            End If
        End If
    Else

        strsqlx = "SELECT ConvertIdPegawaiToShift.IdPegawai,DataPegawai.NamaLengkap,ConvertIdPegawaiToShift.IdShift From ConvertIdPegawaiToShift INNER JOIN DataPegawai ON ConvertIdPegawaiToShift.IdPegawai = DataPegawai.IdPegawai WHERE (NamaLengkap = '" & lvjadwalkerja.SelectedItem & "')"
        Set rsx = Nothing
        Call msubRecFO(rsx, strsqlx)

        If rsx.EOF = True Then Exit Sub
        If rsx(2).Value = "02" Then

            'Dipatok Kalo pegawai Non Shift Selalu pagi
            strSQL = "SELECT TOP 1 Shift FROM V_JadwalKerjaNew WHERE" & _
            " Nama='" & Item.Text & "'" & _
            " AND Shift = 'Pagi'"

        Else

            strSQL = "SELECT TOP 2 Shift FROM V_JadwalKerjaNew WHERE" & _
            " Nama='" & Item.Text & "'" & _
            " AND MONTH(Tanggal)='" & Me.DTPicker1.Month - 1 & "'" & _
            " AND YEAR(Tanggal)='" & Me.DTPicker1.Year & "'" & _
            " ORDER BY Tanggal DESC"

        End If

        Call msubRecFO(rs, strSQL)
        intBanyakShiftTerakhir = 1
        i = 1
        While Not rs.EOF
            strTempShift = rs.Fields.Item("Shift").Value
            If strTempShift = strShiftTerakhir Then
                intBanyakShiftTerakhir = intBanyakShiftTerakhir + 1
            End If
            If i = 1 Then strShiftTerakhir = strTempShift
            i = i + 1
            rs.MoveNext
        Wend
        Item.SubItems(1) = strShiftTerakhir
        If strShiftTerakhir <> "" Then
            If rsx(2).Value = "02" Then
            Else
                Item.SubItems(2) = CStr(intBanyakShiftTerakhir)
            End If
        End If
        
        '//yayang.agus 2014-08-25
        Set rs = Nothing
        strSQL = "SELECT * FROM V_JadwalKerjaNew" & _
                " WHERE Ruangan='" & dcTempatBertugas.Text & "'" & _
                " and Nama='" & Item.Text & "'" & _
                " AND MONTH(Tanggal)='" & Me.DTPicker1.Month & "'" & _
                " AND YEAR(Tanggal)='" & Me.DTPicker1.Year & "'" & _
                " ORDER BY Tanggal"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount = 0 Then Exit Sub
        If fgJadwal.TextMatrix(fgJadwal.Rows - 1, 0) <> "" Then fgJadwal.Rows = fgJadwal.Rows + 1
        fgJadwal.TextMatrix(fgJadwal.Rows - 1, 0) = Item.Text
        Call subIsiFgJadwalKerja
        '//
    End If
End Sub

Private Sub lvjadwalkerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.lvjadwalkerja.ListItems.Count > 0 Then
        If Me.lvjadwalkerja.SelectedItem.Checked Then
            Me.frIsiList.Visible = True
            Me.dcShiftKerja.Text = ""
            Me.dcShiftKerja.SetFocus
        End If
        
    End If
End Sub

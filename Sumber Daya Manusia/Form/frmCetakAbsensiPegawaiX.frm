VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakAbsensiPegawaiX 
   Caption         =   "Cetak Absensi Pegawai"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4710
   Icon            =   "frmCetakAbsensiPegawaiX.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   4710
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakAbsensiPegawaiX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptCetakAbsensi As CrLapAbsensiBulanTahun

Private Sub DateSystem()

    Dim nmonth As Integer
    Dim nLastDay As Integer
    Dim nmodRemainder As Integer
    Dim calDate As Date
    
    calDate = glDateCetak
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
    Dim firstMonthYear As Date
    Dim Datetgl As Date
    Dim iLoop As Integer
    Dim lcsRet As String
    Dim blnAwalBulan As Boolean
    
    blnAwalBulan = True
    firstMonthYear = CDate("01/" & str(Month(calDate)) & "/" & str(Year(calDate)))
    
    For iLoop = 1 To nLastDay
        If blnAwalBulan Then
            Datetgl = firstMonthYear
            blnAwalBulan = False
        Else
            Datetgl = DateSerial(Format(Datetgl, "yyyy"), Format(Datetgl, "MM"), Val(Format(Datetgl, "dd")) + 1)
        End If
        
        strTanggal = Day(Datetgl)
        strBulan = Month(Datetgl)
        strHari = WeekdayName(Weekday(Datetgl), , vbSunday)
        
        Select Case strHari
            Case "Sabtu"
                If Len(strTanggal) = 1 Then
                    lcsRet = "0" & strTanggal
                Else
                    lcsRet = strTanggal
                End If
            Case "Minggu"
                If Len(strTanggal) = 1 Then
                    lcsRet = "0" & strTanggal
                Else
                    lcsRet = strTanggal
                End If
        End Select
        
        Dim ObjRs As New ADODB.recordset
        
        strSQL = "SELECT Tanggal, [Hari Libur] FROM v_tanggal" & _
                 " WHERE DAY(Tanggal)='" & Day(Datetgl) & "'" & _
                 " AND MONTH(Tanggal)='" & Month(Datetgl) & "'" & _
                 " AND YEAR(Tanggal)='" & Year(Datetgl) & "'" & _
                 " ORDER BY Tanggal"
         
         ObjRs.Open strSQL, dbConn, 3, 2
            If Not ObjRs.EOF Then
                If Not IsNull(ObjRs.Fields.Item("Hari Libur").Value) Then
                    If Len(strTanggal) = 1 Then
                        lcsRet = "0" & strTanggal
                    Else
                        lcsRet = strTanggal
                    End If
                End If
            End If
         ObjRs.Close
         
         If lcsRet <> "" Then
             Select Case lcsRet
                Case "01"
                    rptCetakAbsensi.Tgl1.TextColor = vbRed
                Case "02"
                    rptCetakAbsensi.Tgl2.TextColor = vbRed
                Case "03"
                    rptCetakAbsensi.Tgl3.TextColor = vbRed
                Case "04"
                    rptCetakAbsensi.Tgl4.TextColor = vbRed
                Case "05"
                    rptCetakAbsensi.Tgl5.TextColor = vbRed
                Case "06"
                    rptCetakAbsensi.Tgl6.TextColor = vbRed
                Case "07"
                    rptCetakAbsensi.Tgl7.TextColor = vbRed
                Case "08"
                    rptCetakAbsensi.Tgl8.TextColor = vbRed
                Case "09"
                    rptCetakAbsensi.Tgl9.TextColor = vbRed
                Case 10
                    rptCetakAbsensi.Tgl10.TextColor = vbRed
                Case 11
                    rptCetakAbsensi.Tgl11.TextColor = vbRed
                Case 12
                    rptCetakAbsensi.Tgl12.TextColor = vbRed
                Case 13
                    rptCetakAbsensi.Tgl13.TextColor = vbRed
                Case 14
                    rptCetakAbsensi.Tgl14.TextColor = vbRed
                Case 15
                    rptCetakAbsensi.Tgl15.TextColor = vbRed
                Case 16
                    rptCetakAbsensi.Tgl16.TextColor = vbRed
                Case 17
                    rptCetakAbsensi.Tgl17.TextColor = vbRed
                Case 18
                    rptCetakAbsensi.Tgl18.TextColor = vbRed
                Case 19
                    rptCetakAbsensi.Tgl19.TextColor = vbRed
                Case 20
                    rptCetakAbsensi.Tgl20.TextColor = vbRed
                Case 21
                    rptCetakAbsensi.Tgl21.TextColor = vbRed
                Case 22
                    rptCetakAbsensi.Tgl22.TextColor = vbRed
                Case 23
                    rptCetakAbsensi.Tgl23.TextColor = vbRed
                Case 24
                    rptCetakAbsensi.Tgl24.TextColor = vbRed
                Case 25
                    rptCetakAbsensi.Tgl25.TextColor = vbRed
                Case 26
                    rptCetakAbsensi.Tgl26.TextColor = vbRed
                Case 27
                    rptCetakAbsensi.Tgl27.TextColor = vbRed
                Case 28
                    rptCetakAbsensi.Tgl28.TextColor = vbRed
                Case 29
                    rptCetakAbsensi.Tgl29.TextColor = vbRed
                Case 30
                    rptCetakAbsensi.Tgl30.TextColor = vbRed
                Case 31
                    rptCetakAbsensi.Tgl31.TextColor = vbRed
            End Select
        End If
        
    Next iLoop
    
End Sub

Private Sub Form_Load()
    
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    
    Set rptCetakAbsensi = New CrLapAbsensiBulanTahun
    
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    With rptCetakAbsensi
        .Database.AddADOCommand dbConn, dbcmd
        .usNamaPegawai.SetUnboundFieldSource ("{Ado.NamaPegawai}")
        .UnboundDateTime1.SetUnboundFieldSource ("{Ado.TotalAbsensi}")
        .txtRuangan.SetText pubStrRuangan
        .txtAlamat2.SetText strWebsite & " " & strEmail
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txttanggal.SetText pubStrPeriode
        
        Call DateSystem
        
        If subTanggalTerakhir = 30 Then
            .Field35.Suppress = True
            .Tgl31.Suppress = True
            .ltotal.Right = 14280
            .txttotal.Left = 14500
            .FTotal.Left = 14445
            .grand31.Suppress = True
            .ljudul2.Right = 14280
            .FGrandtotal.Left = 14445
        End If
        If subTanggalTerakhir = 28 Then
            .Field33.Suppress = True
            .Field34.Suppress = True
            .Tgl29.Suppress = True
            .Tgl30.Suppress = True
            .Tgl31.Suppress = True
            .Box1.Right = 14280
            .ltotal.Right = 13575
            .ltotal.Left = 13575
            .l28.Suppress = True
            .l29.Suppress = True
            .l30.Suppress = True
            .ljudul.Right = 14280
            .grand29.Suppress = True
            .grand30.Suppress = True
            .grand31.Suppress = True
            .Field35.Suppress = True
            .FTotal.Left = 13600
            .ljudul2.Right = 13575
            .lbawah.Right = 14280
            .txttotal.Left = 13680
            .FGrandtotal.Left = 13600
        End If
        
    End With
    CRViewer1.ReportSource = rptCetakAbsensi
    
    With CRViewer1
        .EnableGroupTree = False
        .EnableExportButton = True
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    With CRViewer1
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakAbsensiPegawaiX = Nothing
End Sub





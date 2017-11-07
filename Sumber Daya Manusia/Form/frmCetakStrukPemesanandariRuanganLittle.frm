VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakStrukPemesanandariRuanganLittle 
   Caption         =   "Medifrst2000 - Struk Pemesanan dari Ruangan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   Icon            =   "frmCetakStrukPemesanandariRuanganLittle.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frmCetakStrukPemesanandariRuanganLittle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crStrukPemesanandariRuanganLittle

Private Sub Form_Load()
    On Error GoTo errLoad
    Dim adocomd As New ADODB.Command
    Dim strsqlx As String
    Dim strKdAsal As String

    Set adocomd = Nothing
    Set frmCetakStrukPemesanandariRuangan = Nothing
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Call openConnection

    adocomd.ActiveConnection = dbConn
    If frmPemesananBarang.dcStatusBarang.BoundText = "02" Then
        strsqlx = "Select KdAsal From DetailorderRuangan where NoOrder Like '%" & mstrNoOrder & "%'"

        Call msubRecFO(rs, strsqlx)
        If rs.EOF = True Then
            strKdAsal = ""
        ElseIf IsNull(rs(0).Value) Then
            strKdAsal = ""
        Else
            If rs.RecordCount <> 1 Then
                rs.MoveNext
                strKdAsal = rs(0).Value
            Else
                strKdAsal = rs(0).Value
            End If
        End If

        adocomd.CommandText = "SELECT * from V_StrukOrderRuanganCetakM WHERE NoOrder = '" & mstrNoOrder & "' and KdAsal Like '%" & strKdAsal & "%'"
        Call msubRecFO(dbRst, "SELECT NoOrder, TglOrder, RuanganPemesan, RuanganTujuan FROM  V_StrukOrderInformasiMedisRuangan WHERE NoOrder = '" & mstrNoOrder & "'")
    ElseIf frmPemesananBarang.dcStatusBarang.BoundText = "01" Then
        adocomd.CommandText = "SELECT * from V_StrukOrderRuanganCetakNM WHERE NoOrder = '" & mstrNoOrder & "' and KdAsal Like '%" & strKdAsal & "%'"
        Call msubRecFO(dbRst, "SELECT NoOrder, TglOrder, RuanganPemesan, RuanganTujuan FROM  V_StrukOrderInformasiNonMedisRuangan WHERE NoOrder = '" & mstrNoOrder & "'")
    End If

    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail

        If dbRst.EOF = False Then
            .txtNoOrder.SetText IIf(IsNull(dbRst("NoOrder")), "", dbRst("NoOrder"))
            .txtTglOrder.SetText IIf(IsNull(dbRst("TglOrder")), "", dbRst("TglOrder"))
            .txtRuanganPemesan.SetText IIf(IsNull(dbRst("RuanganPemesan")), "", dbRst("RuanganPemesan"))
            .txtRuanganTujuan.SetText IIf(IsNull(dbRst("RuanganTujuan")), "", dbRst("RuanganTujuan"))
        End If

        .usNoOrder.SetUnboundFieldSource ("{ado.NoOrder}")
        .usJenisBarang.SetUnboundFieldSource ("{ado.DetailJenisBarang}")
        .usNamaBarang.SetUnboundFieldSource ("{ado.Nama Barang}")
        .unJmlOrder.SetUnboundFieldSource ("{ado.JmlOrder}")
        .usAsalBarang.SetUnboundFieldSource ("{ado.AsalBarang}")
    End With
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
errLoad:
    Screen.MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakStrukPemesanandariRuanganLittle = Nothing
End Sub
VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmLoginBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmLoginBackup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   5175
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Lanjutkan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Tag             =   "*"
         Top             =   690
         Width           =   3015
      End
      Begin VB.TextBox txtUserID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Kata Kunci :"
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
         Left            =   720
         TabIndex        =   8
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nama Pemakai :"
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
         Left            =   405
         TabIndex        =   7
         Top             =   420
         Width           =   1290
      End
   End
   Begin VB.TextBox txtServerName 
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDatabaseName 
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   0
      Picture         =   "frmLoginBackup.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5205
   End
End
Attribute VB_Name = "frmLoginBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Dim adoCommand As New ADODB.Command

    'add arief
    strSQL = "Select KdRuangan,NamaRuangan From Ruangan Where KdRuangan = '181'"
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdRuangan = rs("KdRuangan").Value 'dcRuangan.BoundText
        mstrNamaRuangan = rs("NamaRuangan").Value 'dcRuangan.Text
    End If

    Call msubRecFO(rs, "Select KdInstalasi FROM Ruangan WHERE KdRuangan = '" & mstrKdRuangan & "'")
    If rs.EOF = True Then mstrKdInstalasiLogin = "" Else mstrKdInstalasiLogin = rs("KdInstalasi").Value
    'end arief

    Set rs = Nothing
    rs.Open "Select NamaRS,Alamat,KotaKodyaKab,KodePos,Telepon,NamaFileLogoRS from ProfilRS", dbConn, adOpenStatic, adLockReadOnly
    On Error Resume Next
    mstrKdInstalasiNonMedis = "05"

    strNNamaRS = rs(0).Value
    strNAlamatRS = rs(1).Value
    strNKotaRS = rs(2).Value
    strNKodepos = rs(3).Value
    strNTeleponRS = rs(4).Value
    strNamaFileLogoRS = rs(5).Value
    Set rs = Nothing
    strUser = txtUserID.Text
    strPass = txtPassword.Text
    'edit to SQL 2005, encripsi dr SQL 2005
    strQuery = "SELECT IdPegawai, cast(Username as varchar)as Username , cast(Password as varchar)as Password, Status, KdKategoryUser from Login"
    Call msubRecFO(rsLogin, strQuery)
    If rsLogin.EOF Then Exit Sub

    rsLogin.MoveFirst
    Do While rsLogin.EOF = False
        If UCase(strUser) = UCase(rsLogin!username) And UCase(strPass) = UCase(Crypt(rsLogin!Password)) Then
            strIDPegawaiAktif = rsLogin!idpegawai
            strIDPegawai = rsLogin!idpegawai
            If UCase(strUser) = "ADMIN" Then
                blnAdmin = True
            Else
                blnAdmin = False
            End If
            strQuery = "SELECT * FROM LoginAplikasi WHERE IdPegawai = '" & strIDPegawai & "'"
            Set rsLoginApp = Nothing
            With rsLoginApp
                adoCommand.ActiveConnection = dbConn
                adoCommand.CommandText = strQuery
                adoCommand.CommandType = adCmdText
                Set .Source = adoCommand
                .Open
                If rsLoginApp.RecordCount = 0 Then
                    MsgBox "Anda tidak mempunyai akses untuk membuka aplikasi ini", vbCritical, "Aplikasi Error"
                    Exit Sub
                End If
            End With
            rsLoginApp.MoveFirst
            Do While rsLoginApp.EOF = False
                If rsLoginApp!KdAplikasi = "010" Then GoTo UserPermited
                rsLoginApp.MoveNext
            Loop
            MsgBox "Anda tidak mempunyai akses untuk membuka aplikasi ini", vbCritical, "Aplikasi Error"
            Exit Sub

UserPermited:
            strPassEn = Crypt(txtPassword)
            strQuery = "UPDATE Login SET IdPegawai ='" & _
            strIDPegawai & "', UserName ='" & _
            strUser & "',Password ='" & strPassEn & _
            "',Status = '1' WHERE (IdPegawai = '" & strIDPegawai & "')"
            adoCommand.CommandText = strQuery
            adoCommand.CommandType = adCmdText
            adoCommand.Execute

            strNamaHostLocal = Winsock1.LocalHostName

            Call GetIdPegawai
            UserID = noidpegawai
            MDIUtama.Show
            Unload Me
            Exit Sub
        End If
        rsLogin.MoveNext
    Loop
    MsgBox "Anda salah memasukkan username atau password", vbCritical, "Salah user/password"
End Sub

Private Sub Form_Load()
    Dim adoCommand As New ADODB.Command
    strServerName = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Server Name")
    strDatabaseName = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Database Name")
    txtServerName.Text = strServerName
    txtDatabaseName.Text = strDatabaseName
    strServerName = txtServerName.Text
    strDatabaseName = txtDatabaseName.Text
    If txtServerName.Text = "Error" Then
        MsgBox "Tidak ada nama server"
        frmSetServer.Show
        Unload Me
        Exit Sub
    End If
    Set dbConn = Nothing
    openConnection
    If blnError = True Then Exit Sub
    Exit Sub
errLogin:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLogin = Nothing
End Sub

Private Sub Image1_DblClick()
    Dim strPass As String
    Set rs = Nothing
    rs.Open "Select KdRS From ProfilRS", dbConn, adOpenKeyset, adLockOptimistic
    '    SetTimer hwnd, NV_INPUTBOX, 10, AddressOf TimerProc
    '    strPass = InputBox("Masukan administrator password!")
    '    If strPass <> Trim(rs(0).Value) Then Exit Sub
    Unload Me
    frmSetServer.Show
End Sub

Private Sub Picture3_Click()
    frmSetServer.Show
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Dim StrValid As String
    StrValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
    If KeyAscii = 27 Then
        cmdBatal_Click
    ElseIf KeyAscii = vbKeyBack Then
        Exit Sub
    ElseIf KeyAscii = vbKeyDelete Then
        Exit Sub
    End If
    If InStr(StrValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeySpace Then
        KeyAscii = 0
    End If
    cmdOk.Default = True
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    Dim StrValid As String
    StrValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
    If KeyAscii = 27 Then
        cmdBatal_Click
    ElseIf KeyAscii = 13 Then
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If InStr(StrValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeySpace Then
        KeyAscii = 0
    End If
End Sub


VERSION 5.00
Begin VB.Form frmSetPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Setting Printer"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmSetPrinter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5190
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   11220
      TabIndex        =   16
      Top             =   3840
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Setting Printer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   5190
      Begin VB.ComboBox cboDuplexing 
         Height          =   315
         Left            =   1875
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1125
         Width           =   3165
      End
      Begin VB.ComboBox cboJnsDriver 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   3150
      End
      Begin VB.ComboBox cbojnsPrinter 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   255
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Duplexing"
         Height          =   210
         Left            =   795
         TabIndex        =   15
         Top             =   1155
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Driver"
         Height          =   210
         Left            =   780
         TabIndex        =   14
         Top             =   750
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Printer"
         Height          =   210
         Left            =   795
         TabIndex        =   13
         Top             =   330
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Setting Kertas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   5205
      Begin VB.ComboBox cboUkuranKertas 
         Height          =   315
         Left            =   2565
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   2505
      End
      Begin VB.Frame Frame3 
         Caption         =   "Orientasi Kertas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1335
         TabIndex        =   4
         Top             =   690
         Width           =   3735
         Begin VB.OptionButton OptOrien 
            Caption         =   "Portrait"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1545
            TabIndex        =   6
            Top             =   210
            Width           =   945
         End
         Begin VB.OptionButton OptOrien 
            Caption         =   "Landscape"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2490
            TabIndex        =   5
            Top             =   225
            Width           =   1170
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ukuran Kertas"
         Height          =   210
         Left            =   1350
         TabIndex        =   8
         Top             =   345
         Width           =   1140
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   4920
      Width           =   5190
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2760
         TabIndex        =   2
         Top             =   255
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         TabIndex        =   1
         Top             =   255
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   0
      Picture         =   "frmSetPrinter.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5205
   End
End
Attribute VB_Name = "frmSetPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer

Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cbojnsPrinter.Text = "" Then
        MsgBox "Plih dulu dong Nama Printernya !"
        cbojnsPrinter.SetFocus
        Exit Sub
    ElseIf cboJnsDriver.Text = "" Then
        MsgBox "Pilih dulu dong Driver Printernya !"
        cboJnsDriver.SetFocus
        Exit Sub
    ElseIf cboDuplexing.Text = "" Then
        MsgBox "Pilih dulu dong Jenis Duplexingnya !"
        cboDuplexing.SetFocus
        Exit Sub
    ElseIf cboUkuranKertas.Text = "" Then
        MsgBox "Pilih dulu dong Ukuran Kertasnya !"
        cboUkuranKertas.SetFocus
        Exit Sub
    ElseIf OptOrien(0).Value = False And OptOrien(1).Value = False Then
        MsgBox "Pilih dulu dong Orientasi Kertasnya !"
        OptOrien(0).SetFocus
        Exit Sub
    End If
    sPrinter = cbojnsPrinter.Text
    sDriver = cboJnsDriver.Text
    sDuplex = cboDuplexing.ItemData(cboDuplexing.ListIndex)
    sUkuranKertas = cboUkuranKertas.ItemData(cboUkuranKertas.ListIndex)
    If OptOrien(0).Value = True Then
        sOrientasKertas = crPortrait
    Else
        sOrientasKertas = crLandscape
    End If
    If Text1.Text = "" Then
        OrienKertas = sOrientasKertas
        Text1.ToolTipText = OrienKertas
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    For Each prn In Printers
        cbojnsPrinter.AddItem prn.DeviceName
        cboJnsDriver.AddItem prn.DriverName
    Next
    Addcbo cboUkuranKertas, "Default", crDefaultPaperSize
    Addcbo cboUkuranKertas, "Letter", crPaperLetter
    Addcbo cboUkuranKertas, "Small Letter", crPaperLetterSmall
    Addcbo cboUkuranKertas, "Legal", crPaperLegal
    Addcbo cboUkuranKertas, "10x14", crPaper10x14
    Addcbo cboUkuranKertas, "11x17", crPaper11x17
    Addcbo cboUkuranKertas, "A3", crPaperA3
    Addcbo cboUkuranKertas, "A4", crPaperA4
    Addcbo cboUkuranKertas, "A4 Small", crPaperA4Small
    Addcbo cboUkuranKertas, "A5", crPaperA5
    Addcbo cboUkuranKertas, "B4", crPaperB4
    Addcbo cboUkuranKertas, "B5", crPaperB5
    Addcbo cboUkuranKertas, "C Sheet", crPaperCsheet
    Addcbo cboUkuranKertas, "D Sheet", crPaperDsheet
    Addcbo cboUkuranKertas, "Envelope 9", crPaperEnvelope9
    Addcbo cboUkuranKertas, "Envelope 10", crPaperEnvelope10
    Addcbo cboUkuranKertas, "Envelope 11", crPaperEnvelope11
    Addcbo cboUkuranKertas, "Envelope 12", crPaperEnvelope12
    Addcbo cboUkuranKertas, "Envelope 14", crPaperEnvelope14
    Addcbo cboUkuranKertas, "Envelope B4", crPaperEnvelopeB4
    Addcbo cboUkuranKertas, "Envelope B5", crPaperEnvelopeB5
    Addcbo cboUkuranKertas, "Envelope B6", crPaperEnvelopeB6
    Addcbo cboUkuranKertas, "Envelope C3", crPaperEnvelopeC3
    Addcbo cboUkuranKertas, "Envelope C4", crPaperEnvelopeC4
    Addcbo cboUkuranKertas, "Envelope C5", crPaperEnvelopeC5
    Addcbo cboUkuranKertas, "Envelope C6", crPaperEnvelopeC6
    Addcbo cboUkuranKertas, "Envelope C65", crPaperEnvelopeC65
    Addcbo cboUkuranKertas, "Envelope DL", crPaperEnvelopeDL
    Addcbo cboUkuranKertas, "Envelope Italy", crPaperEnvelopeItaly
    Addcbo cboUkuranKertas, "Envelope Monarch", crPaperEnvelopeMonarch
    Addcbo cboUkuranKertas, "Envelope Personal", crPaperEnvelopePersonal
    Addcbo cboUkuranKertas, "E Sheet", crPaperEsheet
    Addcbo cboUkuranKertas, "Executive", crPaperExecutive
    Addcbo cboUkuranKertas, "Fanfold Legal German", crPaperFanfoldLegalGerman
    Addcbo cboUkuranKertas, "Fanfold Standard German", crPaperFanfoldStdGerman
    Addcbo cboUkuranKertas, "Fanfold US", crPaperFanfoldUS
    Addcbo cboUkuranKertas, "Folio", crPaperFolio
    Addcbo cboUkuranKertas, "Ledger", crPaperLedger
    Addcbo cboUkuranKertas, "Note", crPaperNote
    Addcbo cboUkuranKertas, "Quarto", crPaperQuarto
    Addcbo cboUkuranKertas, "Statement", crPaperStatement
    Addcbo cboUkuranKertas, "Tabloid", crPaperTabloid

    Addcbo cboDuplexing, "Default", crPRDPDefault
    Addcbo cboDuplexing, "Simplex", crPRDPSimplex
    Addcbo cboDuplexing, "Horizontal", crPRDPHorizontal
    Addcbo cboDuplexing, "Vertical", crPRDPVertical
End Sub

Private Sub Addcbo(cbo As ComboBox, Name As String, Index As Integer)
    cbo.AddItem Name
    cbo.ItemData(cbo.NewIndex) = Index
End Sub

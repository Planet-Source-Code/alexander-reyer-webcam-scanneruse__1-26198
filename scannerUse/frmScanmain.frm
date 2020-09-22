VERSION 5.00
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmScanmain 
   Caption         =   "Use existing Scanning-Device"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows-Standard
   Begin ScanLibCtl.ImgScan scnControl 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   3307
      _ExtentY        =   2778
      _StockProps     =   0
      PageType        =   6
      CompressionType =   6
      CompressionInfo =   64
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5940
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   5400
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.OptionButton optPicformat 
      Caption         =   "Option2"
      Height          =   195
      Index           =   1
      Left            =   1860
      TabIndex        =   6
      Top             =   2400
      Width           =   195
   End
   Begin VB.OptionButton optPicformat 
      Caption         =   "Option1"
      Height          =   195
      Index           =   0
      Left            =   1860
      TabIndex        =   5
      Top             =   2220
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show scanned Picture"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   2280
      Width           =   1755
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Use a Scanning-device"
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   2280
      Width           =   2115
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
   Begin VB.DirListBox Dirlistbox 
      Height          =   1890
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   4575
   End
   Begin VB.TextBox txtPath 
      Height          =   555
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmScanmain.frx":0000
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "TIFF"
      Height          =   195
      Left            =   2100
      TabIndex        =   8
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "BMP"
      Height          =   195
      Left            =   2100
      TabIndex        =   7
      Top             =   2220
      Width           =   435
   End
End
Attribute VB_Name = "frmScanmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ~~~~~~~~~~~~~~~~~~~~
'   2001 by Alexander Reyer
'  ~~~~~~~~~~~~~~~~~~~~
Option Explicit

    Dim buffer As String
    Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdScan_Click()
    scan_main
End Sub

Private Sub cmdShow_Click()
    Dim ending As String
    Dim filename As String
    
On Error GoTo fehler
    With scnControl
        Select Case .FileType
            Case BMP_Bitmap
                ending = ".bmp"
            Case TIFF
                ending = ".tif"
        End Select
    End With
    
    filename = txtPath.Text + "\ScannedPicture" + ending
            ShellExecute Me.hwnd, "Open", _
                                    filename, _
                                    vbNullString, vbNullString, vbNormalFocus
        
        Exit Sub
fehler:
    MsgBox Err.Description + " in " + Err.Source + " " + Err.HelpContext
    
End Sub

Private Sub Dirlistbox_Change()
    With Dirlistbox
        txtPath.Text = .List(.ListIndex)
    End With
End Sub

Private Sub Drive1_Change()
    Dirlistbox.Path = Drive1.Drive
End Sub


Private Sub Form_Load()

    Drive1.Drive = App.Path
    Dirlistbox.Path = App.Path
    txtPath.Text = App.Path
    optPicformat_Click (0)
End Sub

Public Sub scan_main()
    Dim scanvalues As ScanToConstants
    
    With scnControl
        If Not .ScannerAvailable Then
            MsgBox "no Scanner-device available", vbOKCancel
            If vbCancel Then .AboutBox
        Else
            .AboutBox
            .ScanTo = DisplayAndFile
            .Image = txtPath.Text + "\ScannedPicture"
            .SetPageTypeCompressionOpts _
                                    BestDisplay, TrueColor24bitRGB, _
                                    JPEGCompression, JPEGHighHigh
            
'            .FileType = BMP_Bitmap

            .ShowSelectScanner
            .OpenScanner
            .StartScan
        End If
     End With

End Sub


Private Sub optPicformat_Click(Index As Integer)
    Select Case Index
        Case 0
            scnControl.FileType = BMP_Bitmap
        Case 1
            scnControl.FileType = TIFF
    End Select
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Comparer"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkGenReport 
      Caption         =   "Generate report? (may take more time in some cases)"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   1320
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   5760
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open File"
      Filter          =   "All Files (*.*)|*.*|"
      Orientation     =   2
   End
   Begin VB.PictureBox picRed 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   9
      Top             =   1920
      Width           =   255
   End
   Begin VB.PictureBox picBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   8
      Top             =   1920
      Width           =   255
   End
   Begin VB.PictureBox picResult 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   7
      Top             =   2280
      Width           =   6060
   End
   Begin VB.CommandButton btnOpen2 
      Caption         =   "Open"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton btnOpen1 
      Caption         =   "Open"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtFile2Path 
      Height          =   285
      Left            =   720
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox txtFile1Path 
      Height          =   285
      Left            =   720
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton btnCompare 
      Caption         =   "Compare"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Caption         =   "Staus: Ready"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   6015
   End
   Begin VB.Label lbl3 
      Caption         =   "Comparing Result:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Difference in files"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Equal in both files"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   408
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Label lbl2 
      Caption         =   "FIle 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lbl1 
      Caption         =   "Fle 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT 'Da rect struct
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long 'API used for drawing

Private redBrush As Long
Private blueBrush As Long 'The color brushes

Private arrDiffArray(100) As Byte 'Array of bits representing grafical difference (100 members for each percent)
Private bSizeDiff As Boolean 'this var vill be set by the compare function if files are different in size

Private Sub btnCompare_Click()

If (Trim(txtFile1Path.Text) = vbNullString) Or (Trim(txtFile2Path.Text) = vbNullString) Then 'check for bad user input
    MsgBox "Please input file names of both files.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If

If txtFile1Path.Text = txtFile2Path.Text Then 'check for bad user input
    MsgBox "You are trying to compare a file to it self. This will not yield results.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If

    lblStatus.Caption = "Status: Comparing... Please wait"
    
    txtFile1Path.Enabled = False 'Disable user controls so the user can not mess something up like pressing compare again
    txtFile2Path.Enabled = False
    btnOpen1.Enabled = False
    btnOpen2.Enabled = False
    btnCompare.Enabled = False
    chkGenReport.Enabled = False
    

    If modComparer.Compare(txtFile1Path, txtFile2Path, chkGenReport, arrDiffArray, bSizeDiff) = False Then
        MsgBox "File comparing has failed. Maybe one of the files is missing or can not be opened.", vbOKOnly + vbCritical, "Error"
        lblStatus.Caption = "Status: Error comparing files"
    Else
        ShowGraficalDifference
        lblStatus.Caption = "Status: Finished"
        If bSizeDiff = True Then lblStatus.Caption = lblStatus.Caption & ". Note: Files are different in size"
    End If
    
    txtFile1Path.Enabled = True 'Enable user controls again
    txtFile2Path.Enabled = True
    btnOpen1.Enabled = True
    btnOpen2.Enabled = True
    btnCompare.Enabled = True
    chkGenReport.Enabled = True
    
    picResult.Refresh 'Refresh the graphics representation so its visible

End Sub

Private Sub btnOpen1_Click()

    On Error GoTo CancelPressed
    
        cdFile.ShowOpen 'open first file
        
        txtFile1Path.Text = cdFile.FileName
        
CancelPressed:

End Sub

Private Sub btnOpen2_Click()

    On Error GoTo CancelPressed
    
        cdFile.ShowOpen 'open second file
        
        txtFile2Path.Text = cdFile.FileName
        
CancelPressed:

End Sub

Private Sub txtFile1Path_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    txtFile1Path.Text = Data.Files.Item(1) 'User draged something
    
End Sub

Private Sub txtFile2Path_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    txtFile2Path.Text = Data.Files.Item(1)
    
End Sub

Function ShowGraficalDifference()

    Dim r As RECT 'Rect to hold drawing position
    Dim n As Long 'Used for counting the loop
    
    redBrush = CreateSolidBrush(255) '255 = red color
    blueBrush = CreateSolidBrush(16711680) 'a shade of blue
    
    For n = 0 To 100
    
        r.Left = n * 4 'calculate rect based on array position
        r.Top = 0
        r.Right = r.Left + 4
        r.Bottom = 33
        
        If arrDiffArray(n) = 0 Then 'Use blue
            FillRect picResult.hdc, r, blueBrush
        Else 'Use red
            FillRect picResult.hdc, r, redBrush
        End If
    Next
    
    DeleteObject redBrush 'Safe clear of objects
    DeleteObject blueBrush
    
End Function

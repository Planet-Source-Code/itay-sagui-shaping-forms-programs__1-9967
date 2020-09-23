VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Shape"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Form"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Last"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   1320
      ScaleHeight     =   5715
      ScaleWidth      =   5235
      TabIndex        =   3
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Form"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoadPic 
      Caption         =   "Load Picture"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Exit the program
Private Sub cmdExit_Click()
    Unload Me
    Unload Form2
End Sub

' Load an array of points
' this is the sub you need to copy to your program
' in order to load an array of points you've created
' using this program
Private Sub cmdLoad_Click()
Dim t As POINTAPI
    With CDialog
' setup the open dialog box
        .DialogTitle = "Open form's shape"
        .Filter = "Form's shape|*.fsp"
        .ShowOpen
        If .FileName <> "" Then ' if filename entered then
            Open .FileName For Binary As #1 ' open filename
            Get #1, , t ' get first empty point
            While Not EOF(1)
                Get #1, , t 'get point
                ReDim Preserve Dots(UBound(Dots) + 1) 'redim array of points
                Dots(UBound(Dots)).X = t.X ' insert point into array
                Dots(UBound(Dots)).Y = t.Y
                List1.AddItem t.X & ", " & t.Y ' add point to list
            Wend
            Close #1
            ReDim Preserve Dots(UBound(Dots) - 1) ' remove last empty point
            List1.RemoveItem List1.ListCount - 1
        End If
    End With
End Sub

' load an image
Private Sub cmdLoadPic_Click()
    With CDialog
        .DialogTitle = "Load a picture"
        .Filter = "All Files|*.*|Images|*.bmp;*.jpg"
        .FilterIndex = 1
        .ShowOpen
        If .FileName <> "" Then
            Picture1 = LoadPicture(.FileName)
        End If
    End With
End Sub

' reset array of points
Private Sub cmdNew_Click()
    List1.Clear
    ReDim Dots(0)
End Sub

' remove last point
Private Sub cmdRemove_Click()
    List1.RemoveItem List1.ListCount - 1
    ReDim Preserve Dots(UBound(Dots) - 1)
End Sub

' save points into file
Private Sub cmdSave_Click()
    With CDialog
        .DialogTitle = "Save form's shape"
        .Filter = "Form's shape|*.fsp"
        .ShowSave
        If .FileName <> "" Then
            Open .FileName For Binary As #1
            Put #1, , Dots()
            Close #1
        End If
    End With
End Sub

' test the created shape
Private Sub cmdTest_Click()
    If cmdTest.Caption = "Test" Then
        ' testing requires at least 3 points (a triangle)
        If UBound(Dots) > 2 Then
            SetForm Dots(1), Form2 ' set form's shape
            Form2.Picture = Picture1 ' load picture in form
            Form2.Show 'show form
            cmdTest.Caption = "Stop Test"
        Else
            MsgBox "You need to have at least 3 points", vbCritical
        End If
    Else
        ' stop testing
        Form2.Hide
        cmdTest.Caption = "Test"
    End If
End Sub

Private Sub Form_Load()
    Load Form2
    ReDim Dots(0)
End Sub

' add new point
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReDim Preserve Dots(UBound(Dots) + 1) ' add new point at the end of array
    Dots(UBound(Dots)).X = X / Screen.TwipsPerPixelX ' insert point
    Dots(UBound(Dots)).Y = Y / Screen.TwipsPerPixelY
    List1.AddItem Dots(UBound(Dots)).X & ", " & Dots(UBound(Dots)).Y
End Sub


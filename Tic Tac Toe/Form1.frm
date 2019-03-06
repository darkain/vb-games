VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   1695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&New"
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.Line Line4 
      X1              =   1560
      X2              =   120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line3 
      X1              =   1560
      X2              =   120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   1080
      X2              =   1080
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   600
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   9
      Left            =   1080
      TabIndex        =   8
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   8
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   6
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   5
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   3
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'started at 9:50 PM
'finished as 11:00 PM
Option Explicit

Sub ClearBoard()
Dim i As Integer
For i = 1 To 9
  Label1(i).Caption = " "
Next i
End Sub

Sub WinMsg()
MsgBox "You Win", , "Tic Tac Toe"
Call ClearBoard
End Sub

Sub LoseMsg()
MsgBox "You Lose", , "Tic Tac Toe"
Call ClearBoard
End Sub

Sub AI_Level1()
Dim TmpStr As String
Dim i As Integer
For i = 1 To 3
  TmpStr = Label1(i) & Label1(i + 3) & Label1(i + 6)
  If TmpStr = " OO" Then
    Label1(i).Caption = "O"
    Call LoseMsg
    Exit Sub
  End If
  If TmpStr = "O O" Then
    Label1(i + 3).Caption = "O"
    Call LoseMsg
    Exit Sub
  End If
  If TmpStr = "OO " Then
    Label1(i + 6).Caption = "O"
    Call LoseMsg
    Exit Sub
  End If

  TmpStr = Label1(i * 3 - 2) & Label1(i * 3 - 1) & Label1(i * 3)
  If TmpStr = " OO" Then
    Label1(i * 3 - 2).Caption = "O"
    Call LoseMsg
    Exit Sub
  End If
  If TmpStr = "O O" Then
    Label1(i * 3 - 1).Caption = "O"
    Call LoseMsg
    Exit Sub
  End If
  If TmpStr = "OO " Then
    Label1(i * 3).Caption = "O"
    Call LoseMsg
    Exit Sub
  End If
Next i

TmpStr = Label1(1) & Label1(5) & Label1(9)
If TmpStr = " OO" Then
  Label1(1).Caption = "O"
  Call LoseMsg
  Exit Sub
End If
If TmpStr = "O O" Then
  Label1(5).Caption = "O"
  Call LoseMsg
  Exit Sub
End If
If TmpStr = "OO " Then
  Label1(9).Caption = "O"
  Call LoseMsg
  Exit Sub
End If

TmpStr = Label1(7) & Label1(5) & Label1(3)
If TmpStr = " OO" Then
  Label1(7).Caption = "O"
  Call LoseMsg
  Exit Sub
End If
If TmpStr = "O O" Then
  Label1(5).Caption = "O"
  Call LoseMsg
  Exit Sub
End If
If TmpStr = "OO " Then
  Label1(3).Caption = "O"
  Call LoseMsg
  Exit Sub
End If

Call AI_Level2
End Sub

Sub AI_Level2()
Dim TmpStr As String
Dim i As Integer
For i = 1 To 3
  TmpStr = Label1(i) & Label1(i + 3) & Label1(i + 6)
  If TmpStr = " XX" Then
    Label1(i).Caption = "O"
    Exit Sub
  End If
  If TmpStr = "X X" Then
    Label1(i + 3).Caption = "O"
    Exit Sub
  End If
  If TmpStr = "XX " Then
    Label1(i + 6).Caption = "O"
    Exit Sub
  End If

  TmpStr = Label1(i * 3 - 2) & Label1(i * 3 - 1) & Label1(i * 3)
  If TmpStr = " XX" Then
    Label1(i * 3 - 2).Caption = "O"
    Exit Sub
  End If
  If TmpStr = "X X" Then
    Label1(i * 3 - 1).Caption = "O"
    Exit Sub
  End If
  If TmpStr = "XX " Then
    Label1(i * 3).Caption = "O"
    Exit Sub
  End If
Next i

TmpStr = Label1(1) & Label1(5) & Label1(9)
If TmpStr = " XX" Then
  Label1(1).Caption = "O"
  Exit Sub
End If
If TmpStr = "X X" Then
  Label1(5).Caption = "O"
  Exit Sub
End If
If TmpStr = "XX " Then
  Label1(9).Caption = "O"
  Exit Sub
End If

TmpStr = Label1(7) & Label1(5) & Label1(3)
If TmpStr = " XX" Then
  Label1(7).Caption = "O"
  Exit Sub
End If
If TmpStr = "X X" Then
  Label1(5).Caption = "O"
  Exit Sub
End If
If TmpStr = "XX " Then
  Label1(3).Caption = "O"
  Exit Sub
End If

Call AI_Level3
End Sub

Sub AI_Level3()
Dim RndNum As Byte
Dim EndLoop As Boolean

Call CheckCat
Do
  RndNum = Int(Rnd * 9) + 1
  If Label1(RndNum).Caption = " " Then
    Label1(RndNum).Caption = "O"
    EndLoop = True
  End If
Loop Until EndLoop
End Sub

Sub CheckCat()
Dim TmpStr As String
Dim i As Integer
Dim Count As Integer

For i = 1 To 9
  If Label1(i).Caption = " " Then Count = Count + 1
Next i
If Count = 0 Then
  MsgBox "Cat Game", , "Tic Tac Toe"
  Call ClearBoard
End If
End Sub

Sub CheckForWin()
Dim TmpStr As String
Dim i As Integer
For i = 1 To 3
  TmpStr = Label1(i) & Label1(i + 3) & Label1(i + 6)
  If TmpStr = "XXX" Then
    Call WinMsg
    Exit Sub
  End If
  TmpStr = Label1(i * 3 - 2) & Label1(i * 3 - 1) & Label1(i * 3)
  If TmpStr = "XXX" Then
    Call WinMsg
    Exit Sub
  End If
Next i
TmpStr = Label1(1) & Label1(5) & Label1(9)
If TmpStr = "XXX" Then
    Call WinMsg
    Exit Sub
End If
TmpStr = Label1(7) & Label1(5) & Label1(3)
If TmpStr = "XXX" Then
    Call WinMsg
    Exit Sub
End If

End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Call ClearBoard
End Sub

Private Sub Form_DblClick()
Call ClearBoard
End Sub

Private Sub Form_Load()
Dim MyNum As Byte
Call ClearBoard
End Sub

Private Sub Label1_Click(Index As Integer)
If Label1(Index).Caption = " " Then
  Label1(Index).Caption = "X"
  Call CheckForWin
  Call AI_Level1
  Call CheckCat
End If
End Sub

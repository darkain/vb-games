VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Checks Out"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2640
      TabIndex        =   102
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Blank"
      Height          =   375
      Left            =   2640
      TabIndex        =   101
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random"
      Height          =   375
      Left            =   2640
      TabIndex        =   100
      Top             =   120
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   99
      Left            =   2280
      TabIndex        =   99
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   98
      Left            =   2040
      TabIndex        =   98
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   97
      Left            =   1800
      TabIndex        =   97
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   96
      Left            =   1560
      TabIndex        =   96
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   95
      Left            =   1320
      TabIndex        =   95
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   94
      Left            =   1080
      TabIndex        =   94
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   93
      Left            =   840
      TabIndex        =   93
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   92
      Left            =   600
      TabIndex        =   92
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   91
      Left            =   360
      TabIndex        =   91
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   90
      Left            =   120
      TabIndex        =   90
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   89
      Left            =   2280
      TabIndex        =   89
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   88
      Left            =   2040
      TabIndex        =   88
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   87
      Left            =   1800
      TabIndex        =   87
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   86
      Left            =   1560
      TabIndex        =   86
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   85
      Left            =   1320
      TabIndex        =   85
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   84
      Left            =   1080
      TabIndex        =   84
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   83
      Left            =   840
      TabIndex        =   83
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   82
      Left            =   600
      TabIndex        =   82
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   81
      Left            =   360
      TabIndex        =   81
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   80
      Left            =   120
      TabIndex        =   80
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   79
      Left            =   2280
      TabIndex        =   79
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   78
      Left            =   2040
      TabIndex        =   78
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   77
      Left            =   1800
      TabIndex        =   77
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   76
      Left            =   1560
      TabIndex        =   76
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   75
      Left            =   1320
      TabIndex        =   75
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   74
      Left            =   1080
      TabIndex        =   74
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   73
      Left            =   840
      TabIndex        =   73
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   72
      Left            =   600
      TabIndex        =   72
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   71
      Left            =   360
      TabIndex        =   71
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   70
      Left            =   120
      TabIndex        =   70
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   69
      Left            =   2280
      TabIndex        =   69
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   68
      Left            =   2040
      TabIndex        =   68
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   67
      Left            =   1800
      TabIndex        =   67
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   66
      Left            =   1560
      TabIndex        =   66
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   65
      Left            =   1320
      TabIndex        =   65
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   64
      Left            =   1080
      TabIndex        =   64
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   63
      Left            =   840
      TabIndex        =   63
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   62
      Left            =   600
      TabIndex        =   62
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   61
      Left            =   360
      TabIndex        =   61
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   60
      Left            =   120
      TabIndex        =   60
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   59
      Left            =   2280
      TabIndex        =   59
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   58
      Left            =   2040
      TabIndex        =   58
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   57
      Left            =   1800
      TabIndex        =   57
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   56
      Left            =   1560
      TabIndex        =   56
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   55
      Left            =   1320
      TabIndex        =   55
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   54
      Left            =   1080
      TabIndex        =   54
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   53
      Left            =   840
      TabIndex        =   53
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   52
      Left            =   600
      TabIndex        =   52
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   51
      Left            =   360
      TabIndex        =   51
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   50
      Left            =   120
      TabIndex        =   50
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   49
      Left            =   2280
      TabIndex        =   49
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   48
      Left            =   2040
      TabIndex        =   48
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   47
      Left            =   1800
      TabIndex        =   47
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   46
      Left            =   1560
      TabIndex        =   46
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   45
      Left            =   1320
      TabIndex        =   45
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   44
      Left            =   1080
      TabIndex        =   44
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   43
      Left            =   840
      TabIndex        =   43
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   42
      Left            =   600
      TabIndex        =   42
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   41
      Left            =   360
      TabIndex        =   41
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   40
      Left            =   120
      TabIndex        =   40
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   39
      Left            =   2280
      TabIndex        =   39
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   38
      Left            =   2040
      TabIndex        =   38
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   37
      Left            =   1800
      TabIndex        =   37
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   36
      Left            =   1560
      TabIndex        =   36
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   35
      Left            =   1320
      TabIndex        =   35
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   34
      Left            =   1080
      TabIndex        =   34
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   33
      Left            =   840
      TabIndex        =   33
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   32
      Left            =   600
      TabIndex        =   32
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   31
      Left            =   360
      TabIndex        =   31
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   30
      Left            =   120
      TabIndex        =   30
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   29
      Left            =   2280
      TabIndex        =   29
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   28
      Left            =   2040
      TabIndex        =   28
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   27
      Left            =   1800
      TabIndex        =   27
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   26
      Left            =   1560
      TabIndex        =   26
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   25
      Left            =   1320
      TabIndex        =   25
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   24
      Left            =   1080
      TabIndex        =   24
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   23
      Left            =   840
      TabIndex        =   23
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   22
      Left            =   600
      TabIndex        =   22
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   21
      Left            =   360
      TabIndex        =   21
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   19
      Left            =   2280
      TabIndex        =   19
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   18
      Left            =   2040
      TabIndex        =   18
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   17
      Left            =   1800
      TabIndex        =   17
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   16
      Left            =   1560
      TabIndex        =   16
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   15
      Left            =   1320
      TabIndex        =   15
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   14
      Left            =   1080
      TabIndex        =   14
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   13
      Left            =   840
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   12
      Left            =   600
      TabIndex        =   12
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   9
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Clicks"
      Height          =   255
      Left            =   2640
      TabIndex        =   103
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Running As Boolean
Dim NumOfClicks As Integer

Private Sub CheckClear()
Dim i As Integer
Dim Count As Integer
For i = 0 To Check1.Count - 1
  Count = Count + Check1(i).Value
Next i
If Count = 0 Then
  MsgBox "You Win!", , "Checks Out"
  Command2_Click
End If
End Sub

Private Sub Check1_Click(Index As Integer)
If Running Then Exit Sub
Running = True

If Index Mod 10 > 0 Then
  Check1(Index - 1).Value = -(Check1(Index - 1).Value) + 1
End If
If Index Mod 10 < 9 Then
  Check1(Index + 1).Value = -(Check1(Index + 1).Value) + 1
End If
If Int(Index / 10) > 0 Then
  Check1(Index - 10).Value = -(Check1(Index - 10).Value) + 1
End If
If Int(Index / 10) < 9 Then
  Check1(Index + 10).Value = -(Check1(Index + 10).Value) + 1
End If
Running = False

NumOfClicks = NumOfClicks + 1
Label1.Caption = "Clicks " & NumOfClicks
CheckClear
End Sub

Private Sub Command1_Click()
Dim i As Integer
Running = True
For i = 0 To Check1.Count - 1
  Check1(i).Value = Int(Rnd * 2)
Next i
Running = False
NumOfClicks = 0
End Sub

Private Sub Command2_Click()
Dim i As Integer
Running = True
For i = 0 To Check1.Count - 1
  Check1(i).Value = 0
Next i
Running = False
NumOfClicks = 0
End Sub

Private Sub Command3_Click()
End
End Sub


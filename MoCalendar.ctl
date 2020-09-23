VERSION 5.00
Begin VB.UserControl MoCalendar 
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   ScaleHeight     =   3210
   ScaleWidth      =   3705
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   255
      Left            =   480
      TabIndex        =   59
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   255
      Left            =   2880
      TabIndex        =   58
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   41
      Left            =   3000
      TabIndex        =   57
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   40
      Left            =   2520
      TabIndex        =   56
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   39
      Left            =   2040
      TabIndex        =   55
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   38
      Left            =   1560
      TabIndex        =   54
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   37
      Left            =   1080
      TabIndex        =   53
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   36
      Left            =   600
      TabIndex        =   52
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   35
      Left            =   120
      TabIndex        =   51
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   34
      Left            =   3000
      TabIndex        =   50
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   33
      Left            =   2520
      TabIndex        =   49
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   32
      Left            =   2040
      TabIndex        =   48
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   31
      Left            =   1560
      TabIndex        =   47
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   30
      Left            =   1080
      TabIndex        =   46
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   29
      Left            =   600
      TabIndex        =   45
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   28
      Left            =   120
      TabIndex        =   44
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   27
      Left            =   3000
      TabIndex        =   43
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   26
      Left            =   2520
      TabIndex        =   42
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   25
      Left            =   2040
      TabIndex        =   41
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   24
      Left            =   1560
      TabIndex        =   40
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   23
      Left            =   1080
      TabIndex        =   39
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   22
      Left            =   600
      TabIndex        =   38
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   37
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   20
      Left            =   3000
      TabIndex        =   36
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   19
      Left            =   2520
      TabIndex        =   35
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   18
      Left            =   2040
      TabIndex        =   34
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   17
      Left            =   1560
      TabIndex        =   33
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   16
      Left            =   1080
      TabIndex        =   32
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   15
      Left            =   600
      TabIndex        =   31
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   30
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   13
      Left            =   3000
      TabIndex        =   29
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   12
      Left            =   2520
      TabIndex        =   28
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   11
      Left            =   2040
      TabIndex        =   27
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   10
      Left            =   1560
      TabIndex        =   26
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   25
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   24
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   22
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   21
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   20
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   19
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   18
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   17
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Day_Cap 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   13
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   12
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   11
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   10
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   3720
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   3360
      Y1              =   480
      Y2              =   3600
   End
   Begin VB.Label Label2 
      Caption         =   "Sa"
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   7
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Fr"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Th"
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "We"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Tu"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Mo"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Su"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "MoCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Public Event Declarations

Public Event ChangeMonth()
Public Event GotDate()
Public Event ClickDate()


Private Sub Command1_Click()  'Next Month
Month1 = Month1 + 1
If Month1 > 12 Then
    Month1 = 1
    Year1 = Year1 + 1
End If
RaiseEvent ChangeMonth
UpdateCal
End Sub

Private Sub Command2_Click()  'Previous Month
Month1 = Month1 - 1
If Month1 < 1 Then
    Month1 = 12
    Year1 = Year1 - 1
End If
RaiseEvent ChangeMonth
UpdateCal
End Sub

Private Sub Command3_Click()  'Advance Year
Year1 = Year1 + 1
RaiseEvent ChangeMonth
UpdateCal
End Sub

Private Sub Command4_Click()  'Previous Year
Year1 = Year1 - 1
RaiseEvent ChangeMonth
UpdateCal
End Sub

Private Sub Day_Cap_Click(Index As Integer)
If Val(Day_Cap(Index).Caption) > 0 Then  'If a day was clicked
    Day1 = Val(Day_Cap(Index).Caption)   'set current day to the clicked day
    UpdateCal
    RaiseEvent ClickDate                 'tell other programs the day was clicked
End If
End Sub

Private Sub Day_Cap_DblClick(Index As Integer)
If Val(Day_Cap(Index).Caption) > 0 Then  'If a day was clicked (VAL("")=0)
    Day1 = Val(Day_Cap(Index).Caption)   'set current day to the clicked
    UpdateCal
    RaiseEvent GotDate
End If
Debug.Print "DAY=" + Str(Day1)
End Sub

Private Sub UserControl_Initialize()
'Clear MonthView
ClearMonth
'Set Basic Variabled to today
Month1 = Month(Now)
Day1 = Day(Now)
Year1 = Year(Now)
'Launch the subroutine that fills in the month
UpdateCal
End Sub

Private Sub PutDay(dayofweek As Integer, week As Integer, DayLabel As String)
Dim TheNumber As Integer
TheNumber = (7 * (week - 1)) + dayofweek - 1
Day_Cap(TheNumber).Caption = DayLabel  'change label to the day
If DayLabel = Str(Day1) Then           'if today is highlited, then make bold
   Day_Cap(TheNumber).FontBold = True
Else
   Day_Cap(TheNumber).FontBold = False
End If
Day_Cap(TheNumber).Width = 300
End Sub

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
  UserControl.Enabled() = vNewValue
  PropertyChanged "Enabled"
End Property

Public Property Get TheDate() As Date
  Dim MyDate As String
  MyDate = Trim(Str(Month1)) + "-" + Trim(Str(Day1)) + "-" + Trim(Str(Year1))
  TheDate = DateValue(MyDate)   ' send date out
End Property

Public Property Let TheDate(ByVal vNewValue As Date)
  Month1 = Month(vNewValue)
  Day1 = Day(vNewValue)
  Year1 = Year(vNewValue)
  UpdateCal         'let date be set by outside program
End Property

Private Sub UpdateCal()
  ClearMonth        'remove old garbage
  Label1.Caption = MonthName(Month1) + " - " + Str(Year1)
  Dim MyDate As Date
  Dim myweek As Integer
  Dim MaxMonth As Date
  Dim n As Integer
  MyDate = Trim(Str(Month1)) + "-" + Trim(Str(1)) + "-" + Trim(Str(Year1)) ' this is "DATE" type
  Label3(0).Caption = WeekNumber(MyDate)
  Select Case Month1   'find the max number of days that month
    Case 1
        MaxMonth = 31
    Case 2
        MaxMonth = 28
    Case 3
        MaxMonth = 31
    Case 4
        MaxMonth = 30
    Case 5
        MaxMonth = 31
    Case 6
        MaxMonth = 30
    Case 7
        MaxMonth = 31
    Case 8
        MaxMonth = 31
    Case 9
        MaxMonth = 30
    Case 10
        MaxMonth = 31
    Case 11
        MaxMonth = 30
    Case 12
        MaxMonth = 31
End Select
  If Day1 > MaxMonth Then Day1 = MaxMonth  ' so we can't pass back bad vars
  myweek = 1
  For n = 1 To MaxMonth
    MyDate = Trim(Str(Month1)) + "/" + Trim(Str(n)) + "/" & Trim(Str(Year1))
    x = DatePart("w", MyDate, vbSunday, vbFirstFourDays)
    If (x = 1) And (n > 1) Then
        myweek = myweek + 1
    ElseIf x = 2 Then
        Label3(myweek - 1).Caption = WeekNumber(MyDate)  'fill in week number
    End If
    PutDay Val(x), myweek, Str(n)
  Next n
End Sub


Public Sub ClearMonth()
Dim y As Integer
Dim x As Integer
For y = 1 To 6
For x = 1 To 7
  PutDay x, y, " "
Next x, y
Label1.Caption = " "
For x = 0 To 5
    Label3(x).Caption = " "
Next
End Sub

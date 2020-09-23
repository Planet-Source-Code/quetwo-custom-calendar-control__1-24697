Attribute VB_Name = "Module1"
'PicasCAL
'----------------------------------------
'Written by Nick Kwiatkowski
'kwiatk27@egr.msu.edu
'07/01/01
'----------------------------------------


'Global Vars
Global Month1 As Integer
Global Day1 As Integer
Global Year1 As Integer

'This function was given to me by Microsoft per
'Helpdesk number Q200299
'------
'This is the proper way to find WeekNumbers

Function WeekNumber(InDate As Date) As Integer
  Dim DayNo As Integer
  Dim StartDays As Integer
  Dim StopDays As Integer
  Dim StartDay As Integer
  Dim StopDay As Integer
  Dim VNumber As Integer
  Dim ThurFlag As Boolean

  DayNo = Days(InDate)
  StartDay = Weekday(DateSerial(Year(InDate), 1, 1)) - 1
  StopDay = Weekday(DateSerial(Year(InDate), 12, 31)) - 1
  ' Number of days belonging to first calendar week
  StartDays = 7 - (StartDay - 1)
  ' Number of days belonging to last calendar week
  StopDays = 7 - (StopDay - 1)
  ' Test to see if the year will have 53 weeks or not
  If StartDay = 4 Or StopDay = 4 Then ThurFlag = True Else ThurFlag = False
  VNumber = (DayNo - StartDays - 4) / 7
  ' If first week has 4 or more days, it will be calendar week 1
  ' If first week has less than 4 days, it will belong to last year's
  ' last calendar week
  If StartDays >= 4 Then
     WeekNumber = Fix(VNumber) + 2
  Else
     WeekNumber = Fix(VNumber) + 1
  End If
  ' Handle years whose last days will belong to coming year's first
  ' calendar week
  If WeekNumber > 52 And ThurFlag = False Then WeekNumber = 1
  ' Handle years whose first days will belong to the last year's
  ' last calendar week
  If WeekNumber = 0 Then
     WeekNumber = WeekNumber(DateSerial(Year(InDate) - 1, 12, 31))
  End If
End Function

Function Days(DayNo As Date) As String
  Days = DayNo - DateSerial(Year(DayNo), 1, 0)
End Function

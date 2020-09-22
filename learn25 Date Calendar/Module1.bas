Attribute VB_Name = "Module1"
Option Explicit
Function CountDaysInAYear(DateEntered As String) As Integer
Dim DateDay, DateMonth, DateYear, iMonth(12), _
    loop1, TotalDays As Integer

iMonth(1) = 31 'Jan
iMonth(2) = 28 'Feb
iMonth(3) = 31 'Mar
iMonth(4) = 30 'Apr
iMonth(5) = 31 'May
iMonth(6) = 30 'Jun
iMonth(7) = 31 'Jul
iMonth(8) = 31 'Aug
iMonth(9) = 30 'Sep
iMonth(10) = 31 'Oct
iMonth(11) = 30 'Nov
iMonth(12) = 31 'Dec

   DateYear = Year(DateEntered)
   DateMonth = Month(DateEntered)
   DateDay = Day(DateEntered)
'Check if year is leapyear
If DateYear Mod 4 = 0 Then iMonth(2) = 29
TotalDays = 0
For loop1 = 1 To DateMonth
   If loop1 = DateMonth Then TotalDays = TotalDays + DateDay
   If loop1 < DateMonth Then TotalDays = TotalDays + iMonth(loop1)
Next loop1

CountDaysInAYear = TotalDays

End Function
Function CountNumberOfDaysFromJan1Year1ToDec31YearEntered(YearEntered As Long) As Long
    Dim TotalDays, Year As Long
    Dim Days, DaysInAYear, Counter As Integer

    TotalDays = 0
    Counter = 1
    DaysInAYear = 365

    For Year = 1 To YearEntered
        For Days = 1 To DaysInAYear
            TotalDays = TotalDays + 1
        Next Days
        Counter = Counter + 1
        If Counter = 4 Then
            Counter = 0
            DaysInAYear = 366
        Else
            DaysInAYear = 365
        End If
    Next Year

    CountNumberOfDaysFromJan1Year1ToDec31YearEntered = TotalDays

End Function

Function GETDATE(NumberOfDaysFromJanYear1 As Long) As String
    Dim YearCount As Integer
    Dim Days, i, iMonth(12) As Integer
    Dim GET_Day As Long
    Dim strDay As String
    'Get Day
     GET_Day = NumberOfDaysFromJanYear1 Mod 7
     
   Select Case GET_Day
    Case 1: strDay = "Sunday"
    Case 2: strDay = "Monday"
    Case 3: strDay = "Tuesday"
    Case 4: strDay = "Wednesday"
    Case 5: strDay = "Thursday"
    Case 6: strDay = "Friday"
    Case 0: strDay = "Saturday"
   End Select
    
    iMonth(1) = 31 'Jan
    iMonth(2) = 28 'Feb
    iMonth(3) = 31 'Mar
    iMonth(4) = 30 'Apr
    iMonth(5) = 31 'May
    iMonth(6) = 30 'Jun
    iMonth(7) = 31 'Jul
    iMonth(8) = 31 'Aug
    iMonth(9) = 30 'Sep
    iMonth(10) = 31 'Oct
    iMonth(11) = 30 'Nov
    iMonth(12) = 31 'Dec
    
    'Get Year
    YearCount = 1
    Days = 365
    Do While NumberOfDaysFromJanYear1 > Days
       NumberOfDaysFromJanYear1 = NumberOfDaysFromJanYear1 - Days
       If NumberOfDaysFromJanYear1 > 0 Then YearCount = YearCount + 1
       If YearCount Mod 4 = 0 Then
          Days = 366
       Else
          Days = 365
       End If
    Loop
    
    'Get Month and Day
    If YearCount Mod 4 = 0 Then iMonth(2) = 29
    If NumberOfDaysFromJanYear1 <> 0 Then
      i = 1
      Do While NumberOfDaysFromJanYear1 > iMonth(i)
         NumberOfDaysFromJanYear1 = NumberOfDaysFromJanYear1 - iMonth(i)
         i = i + 1
      Loop
    Else
      i = 12
      NumberOfDaysFromJanYear1 = 31
    End If
    
   GETDATE = str(i) & "/" & str(NumberOfDaysFromJanYear1) & "/" & str(YearCount)
    
End Function
Function GetYear(str As String) As Integer
    Dim loop1, Counter As Integer
    Dim tmpstr As String
    
    For loop1 = 1 To Len(str)
        tmpstr = tmpstr & Mid(str, loop1, 1)
        If Mid(str, loop1, 1) = "/" Then tmpstr = ""
    Next
    GetYear = Val(tmpstr)
End Function
Function GetDateNdaysAgoFromReferenceDate(RefDate As String, DaysBack As Long) As String
    Dim Ref As Long
    Dim Year As String
    
    Ref = CountNumberOfDaysFromJan1Year1ToDec31YearEntered(GetYear(Trim(RefDate)) - 1) + CountDaysInAYear(Format(RefDate, "mm/dd/yyyy"))
    Ref = Ref - DaysBack
    GetDateNdaysAgoFromReferenceDate = GETDATE(Ref)
End Function

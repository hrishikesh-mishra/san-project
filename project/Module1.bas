Attribute VB_Name = "VERIFYDATE"
'THIS FUNCTION CHECK THE DATE FOR VALIDITY


Public Function VERIFY_DATE(DAY As Integer, MONTH As String, YEAR As Integer) As Boolean
    
    Dim leapYear As Boolean
    Dim validDate As Boolean
    
    If (YEAR Mod 4 = 0 And YEAR Mod 100 <> 0) Or YEAR Mod 400 = 0 Then
        leapYear = True
    Else
        leapYear = False
    End If
 
    If (MONTH = "January" Or MONTH = "March" Or MONTH = "May" Or _
        MONTH = "July" Or MONTH = "August" Or MONTH = "October" Or _
        MONTH = "December") And (DAY <= 31) Then
             validDate = True
    ElseIf (MONTH = "April" Or MONTH = "June" Or MONTH = "September" _
            Or MONTH = "November") And (DAY <= 30) Then
             validDate = True
    ElseIf (MONTH = "February") And (leapYear) And (DAY <= 29) Then
              validDate = True
    ElseIf (MONTH = "February") And (leapYear = False) And (DAY <= 28) Then
              validDate = True
    Else
              validDate = False
    End If

    VERIFY_DATE = validDate
 
 End Function

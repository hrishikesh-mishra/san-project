Attribute VB_Name = "FILLWITHDATE"
'THIS SUBROUTINE FILL THE DATES IN THE COMBOS IN THE FORM WHERE IT CALL FROM

Public Sub FILLCOMBODATE(dayCombo As ComboBox, monthCombo As ComboBox, yearCombo As ComboBox)
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer

    For i = 1 To 31
        dayCombo.AddItem i
    Next
    
    For j = 1 To 12
        monthCombo.AddItem MonthName(j)
        
    Next

    For K = 1980 To 2100
        yearCombo.AddItem K
    Next

    dayCombo.Text = DAY(Date)
    monthCombo.Text = MonthName(MONTH(Date))
    yearCombo.Text = YEAR(Date)

End Sub

Public Class frmAbsent
    Dim cmdAbs As New SqlClient.SqlCommand
    Dim rdAbs As SqlClient.SqlDataReader
    Dim Today As DateTime
    Dim Yesterday As DateTime
    Dim DtX As DateTime

    Dim DayNum As Integer
    Dim DayNum1 As Integer





    
    Private Sub frmAbsent_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Today = DateTime.Now
        Yesterday = Today.AddDays(-1)
        DtX = Today.AddDays(-1)

        Select Case Dtx.DayOfWeek.ToString
            Case "Sunday"
                DayNum1 = 0
            Case "Monday"
                DayNum1 = 1
            Case "Tuesday"
                DayNum1 = 2
            Case "Wednesday"
                DayNum1 = 3
            Case "Thursday"
                DayNum1 = 4
            Case "Friday"
                DayNum1 = 5
            Case "Saturday"
                DayNum1 = 6
        End Select


        cmdAbs.CommandText = "select tableEmployee.EmpNo,tableEmployee.FirstName,weekendday from tableEmployee where  tableEmployee.EmpNo  " & _
" not  in (select empno from tableEmpAttendance where convert(date,tableEmpAttendance.EntryDate)='" & X & "') and tableEmployee.EmpStatus=1"









        'Apply Absent or Rest logic
        '----------------------------------------------
        If DayNum = DayNum1 Then
            RuleShort = "R"
        ElseIf DayNum <> DayNum1 And DayOfWeek.Sunday Then
            RuleShort = "R"
        Else
            RuleShort = "A"
        End If
        '----------------------------------------------

















    End Sub
End Class
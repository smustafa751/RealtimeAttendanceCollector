Imports Microsoft.Win32
Imports System.IO
Imports System.Drawing.Imaging

Public Class frmConAtt4

    Public axCZKEM1 As New zkemkeeper.CZKEM


    Dim cmdDev As New SqlClient.SqlCommand
    Dim rdDev As SqlClient.SqlDataReader

    Dim strIP As String
    Dim strPort As String
    Dim RecordFound As String
    Dim MachineName As String

    Dim sEmpID As String
    Dim MakeDate As String



    Dim cmdEmp As New SqlClient.SqlCommand
    Dim rdEmp As SqlClient.SqlDataReader

    Dim ResetCounter As Integer


    Dim cmdSQL As New SqlClient.SqlCommand
    Dim rdSQL As SqlClient.SqlDataReader

    Dim cmdShift As New SqlClient.SqlCommand
    Dim rdShift As SqlClient.SqlDataReader
    Dim ShiftID As Integer
    Dim MustWork As Decimal



    Dim Ds As DateTime
    Dim Ds2 As DateTime

    Dim Rec1 As Integer
    Dim PrevRec As Integer
    Dim LastEntry As Boolean

    Dim WorkHours As Decimal
    Dim TimeDiff As TimeSpan
    Dim Extra As Decimal
    Dim Less As Decimal
    Dim CalHours As Decimal



    Dim cmdWeekEnd As New SqlClient.SqlCommand
    Dim rdWeekEnd As SqlClient.SqlDataReader
    Dim DayNum As Integer
    Dim Dtx As Date

    Dim RuleShort As String

#Region "Communication"
    Private bIsConnected = False 'the boolean value identifies whether the device is connected
    Private iMachineNumber As Integer 'the serial number of the device.After connecting the device ,this value will be changed.


    'If your device supports the TCP/IP communications, you can refer to this.
    'when you are using the tcp/ip communication,you can distinguish different devices by their IP address.
    'Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
    Private Sub ConnectDev()
        If strIP = "" Or strPort = "" Then
            ListBox2.Items.Add("IP and Port cannot be null")
            Return
        End If
        Dim idwErrorCode As Integer
        Cursor = Cursors.WaitCursor
        If button2.Text = "Disconnect" Then
            axCZKEM1.Disconnect()

            RemoveHandler axCZKEM1.OnAttTransactionEx, AddressOf AxCZKEM1_OnAttTransactionEx

            bIsConnected = False
            button2.Text = "Connect"
            Label2.Text = "Current State:Disconnected"
            Cursor = Cursors.Default
            Return
        End If

        bIsConnected = axCZKEM1.Connect_Net(strIP.Trim(), Convert.ToInt32(strPort.Trim()))
        If bIsConnected = True Then
            button2.Text = "Disconnect"
            button2.Refresh()
            Label2.Text = "Current State:Connected to " & RecordFound
            iMachineNumber = 1 'In fact,when you are using the tcp/ip communication,this parameter will be ignored,that is any integer will all right.Here we use 1.

            If axCZKEM1.RegEvent(iMachineNumber, 65535) = True Then 'Here you can register the realtime events that you want to be triggered(the parameters 65535 means registering all)

                AddHandler axCZKEM1.OnAttTransactionEx, AddressOf AxCZKEM1_OnAttTransactionEx

            End If
        Else
            axCZKEM1.GetLastError(idwErrorCode)
            ListBox2.Items.Add("Unable to connect the device,ErrorCode = " & idwErrorCode)
        End If
        Cursor = Cursors.Default


    End Sub


#End Region

#Region "Realtime Events"
    Private Sub AxCZKEM1_OnAttTransactionEx(ByVal sEnrollNumber As String, ByVal iIsInValid As Integer, ByVal iAttState As Integer, ByVal iVerifyMethod As Integer, _
                      ByVal iYear As Integer, ByVal iMonth As Integer, ByVal iDay As Integer, ByVal iHour As Integer, ByVal iMinute As Integer, ByVal iSecond As Integer, ByVal iWorkCode As Integer)
        'ListBox1.Items.Add("RTEvent OnAttTrasactionEx Has been Triggered,Verified OK")
        ListBox1.Items.Add("...UserID:" & sEnrollNumber)
        'ListBox1.Items.Add("...isInvalid:" & iIsInValid.ToString())
        'ListBox1.Items.Add("...attState:" & iAttState.ToString())
        'ListBox1.Items.Add("...VerifyMethod:" & iVerifyMethod.ToString())
        'ListBox1.Items.Add("...Workcode:" & iWorkCode.ToString()) 'the difference between the event OnAttTransaction and OnAttTransactionEx
        ListBox1.Items.Add("...Time:" & iYear.ToString() & "-" & iMonth.ToString() & "-" & iDay.ToString() & " " & iHour.ToString() & ":" & iMinute.ToString() & ":" & iSecond.ToString())

        sEmpID = sEnrollNumber

        MakeDate = iYear.ToString() & "-" + iMonth.ToString() & "-" & iDay.ToString() & " " & iHour.ToString() & ":" & iMinute.ToString() & ":" & iSecond.ToString()
        LoadUserInfo()

    End Sub


#End Region



    Private Sub Load1stMachine()
        Try
            If Con.State = ConnectionState.Closed Then
                DBCON.OpenCon()
            End If
            RecordFound = ""

            cmdDev.CommandText = "select machinename, ipaddress, port from tabledevice where machincenumber=4"
            cmdDev.Connection = Con
            rdDev = cmdDev.ExecuteReader
            While rdDev.Read
                If Not IsDBNull(rdDev(1)) Then
                    RecordFound = rdDev(0)
                    strIP = rdDev(1)
                    strPort = rdDev(2)
                    MachineName = rdDev(0)
                    Me.Text = MachineName & " - " & strIP
                End If
            End While
            rdDev.Close()
            cmdDev.Dispose()
            If RecordFound = "" Then
                ListBox2.Items.Add("Device Not found! for " & MachineName)
                rdDev.Close()
                cmdDev.Dispose()
                Con.Close()

            Else
                rdDev.Close()
                cmdDev.Dispose()
                Con.Close()
                ConnectDev()
            End If


        Catch ex As Exception
            ListBox2.Items.Add("Can not proceed due to : " & ex.ToString)

        End Try
    End Sub








    Private Sub LoadUserInfo()
        Try
            If Con.State = ConnectionState.Closed Then
                DBCON.OpenCon()
            End If
            If sEmpID <> "" Then
                cmdEmp.CommandText = "select firstname + ' ' + lastname as [Name], jobtitle as [Designation], longname as [Department], " & _
                    " img, shiftname as [Shift], empno,shiftid from view_emp_details where empno='" & sEmpID & "'"

                cmdEmp.Connection = Con
                rdEmp = cmdEmp.ExecuteReader

                While rdEmp.Read

                    If rdEmp(5) <> "" Then
                        Label3.Text = "Name: " & rdEmp(0) & vbCrLf & "Designation: " & rdEmp(1) & vbCrLf & "Department: " & rdEmp(2) & vbCrLf & "Shift: " & rdEmp(4)
                        'If Not IsDBNull(rdEmp("img")) Then
                        '    Dim bytBLOBData(rdEmp.GetBytes(3, 0, Nothing, 0, Integer.MaxValue) - 1) As Byte

                        '    rdEmp.GetBytes(3, 0, bytBLOBData, 0, bytBLOBData.Length)
                        '    Dim stmBLOBData As New MemoryStream(bytBLOBData)
                        '    PictureBox1.Image = Image.FromStream(stmBLOBData)
                        'Else
                        '    PictureBox1.Image = Nothing
                        'End If

                    Else
                        Label3.Text = "Data not found for ID: " & sEmpID

                    End If


                    GoReal()


                End While
                ResetCounter = 0
                rdEmp.Close()
                cmdEmp.Dispose()
                'Con.Close()

            Else
                ListBox2.Items.Add("Employee info not found! for " & sEmpID)

            End If

        Catch ex As Exception
            ListBox2.Items.Add("Can not proceed due to : " & ex.ToString)

        End Try
    End Sub


    Private Sub GoReal()
        Dim DayNum1 As Integer

        If LastEntry = True Then LastEntry = False
        Ds = MakeDate
        Ds2 = MakeDate
        Dtx = MakeDate
        'Dtx = Convert.ToDateTime(Dtx.ToString("MM/dd/yyyy"))


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


        'WorkHours = DateDiff(DateInterval.Minute, Ds, Ds2)
        'WorkHours = WorkHours / 60

        If CON1.State = ConnectionState.Closed Then
            DBCON.OpenCon1()
        End If

        Try




            cmdSQL.CommandText = "select count(empno) from tableempattendance where empno='" & sEmpID & "' and convert(date,entrydate)='" & Ds.ToString("MM/dd/yyyy") & "' and outdate is null"
            cmdSQL.Connection = CON1

            rdSQL = cmdSQL.ExecuteReader

            While rdSQL.Read
                Rec1 = rdSQL(0)
            End While
            rdSQL.Close()
            cmdSQL.Dispose()

            If Rec1 = 0 Then
                Ds2 = Ds2.AddDays(-1)
                cmdSQL.CommandText = "select count(empno) from tableempattendance where empno='" & sEmpID & "' and convert(date,entrydate)='" & Ds2.ToString("MM/dd/yyyy") & "' and outdate is null"
                cmdSQL.Connection = CON1

                rdSQL = cmdSQL.ExecuteReader

                While rdSQL.Read
                    PrevRec = rdSQL(0)
                End While
                rdSQL.Close()
                cmdSQL.Dispose()


                If PrevRec > 0 Then
                    'TimeDiff = Ds.Subtract(Ds2)
                    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++
                    'Process Attendance
                    cmdShift.CommandText = "select distinct shiftid,dutyhours from View_ForEmpAttendance where empno='" & sEmpID & "' and statusid=1"
                    cmdShift.Connection = CON1
                    rdShift = cmdShift.ExecuteReader
                    While rdShift.Read
                        If Not IsDBNull(rdShift(0)) Then
                            ShiftID = rdShift(0)
                            MustWork = rdShift(1)
                        Else
                            ShiftID = 0
                        End If

                    End While
                    rdShift.Close()
                    cmdShift.Dispose()
                    Dim cmdMulti As New SqlClient.SqlCommand
                    Dim rdMulti As SqlClient.SqlDataReader
                    Dim LastTime As DateTime
                    Dim TimePassed As Integer

                    cmdMulti.CommandText = "select entrydate from  tableempattendance where empno='" & sEmpID & "' and outdate is null and convert(date,entrydate)='" & Ds2.ToString("MM/dd/yyyy") & "'"
                    cmdMulti.Connection = CON1

                    rdMulti = cmdMulti.ExecuteReader
                    While rdMulti.Read
                        LastTime = rdMulti(0)
                    End While
                    rdMulti.Close()
                    cmdMulti.Dispose()

                    TimeDiff = Ds.Subtract(LastTime)

                    WorkHours = TimeDiff.TotalHours
                    Extra = 0
                    Less = 0
                    If WorkHours > MustWork Then
                        Extra = WorkHours - MustWork
                    ElseIf WorkHours < MustWork Then
                        Less = MustWork - WorkHours
                    Else
                        Extra = 0
                        Less = 0

                    End If


                    '=================================================
                    'Handle RestDay Present as OT


                    cmdWeekEnd.CommandText = "select empno,weekendday from tableemployee where empno='" & sEmpID & "'"
                    cmdWeekEnd.Connection = CON1
                    rdWeekEnd = cmdWeekEnd.ExecuteReader

                    While rdWeekEnd.Read
                        DayNum = rdWeekEnd(1)
                    End While
                    rdWeekEnd.Close()
                    cmdWeekEnd.Dispose()


                    If DayNum = DayNum1 Then
                        RuleShort = "OT"
                    ElseIf DayNum <> DayNum1 And DayOfWeek.Sunday Then
                        RuleShort = "OT"
                    Else
                        RuleShort = "P"
                    End If



                    '==================================================






                    cmdSQL.CommandText = "update tableempattendance set outdate='" & Ds.ToString("MM/dd/yyyy  HH:mm:ss") & "',ruleshort='" & RuleShort & "',recordcomplete=1,machinesn='" & MachineName & "',empshiftid='" & ShiftID & "', workhours=" & WorkHours & ", mustwork=" & MustWork & ",extra1=" & Extra & ",less=" & Less & "  where empno='" & sEmpID & "' and outdate is null and convert(date,entrydate)='" & Ds2.ToString("MM/dd/yyyy") & "'"

                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++


                    'cmdSQL.CommandText = "update tableempattendance set outdate='" & Ds.ToString("MM/dd/yyyy  HH:mm:ss") & "' where empno='" & sEmpID & "' and outdate is null and convert(date,entrydate)='" & Ds2.ToString("MM/dd/yyyy") & "'"
                    cmdSQL.Connection = CON1
                    cmdSQL.ExecuteScalar()
                    LastEntry = True
                End If

                rdSQL.Close()
                cmdSQL.Dispose()

            End If

            Dim cmdDup As New SqlClient.SqlCommand
            Dim rdDup As SqlClient.SqlDataReader
            Dim Duplicate As Integer

            cmdDup.CommandText = "select top 1 count(empno) from tableempattendance where empno='" & sEmpID & "' and convert(date,entrydate)='" & Ds.ToString("MM/dd/yyyy  HH:mm:ss") & "' and convert(date,OutDate )='" & Ds.ToString("MM/dd/yyyy  HH:mm:ss") & "'"
            cmdDup.Connection = CON1
            rdDup = cmdDup.ExecuteReader
            While rdDup.Read
                Duplicate = rdDup(0)

            End While
            rdDup.Close()
            cmdDup.Dispose()

            If LastEntry = False Then
                If Duplicate = 0 Then
                    If Rec1 = 0 And PrevRec = 0 Then
                        cmdSQL.CommandText = "insert into tableempattendance (empno,entrydate)  values ('" & sEmpID & "','" & Ds.ToString("MM/dd/yyyy HH:mm:ss") & "')"
                        cmdSQL.Connection = CON1
                        cmdSQL.ExecuteScalar()
                        cmdSQL.Dispose()
                    Else

                        cmdShift.CommandText = "select distinct shiftid,dutyhours from View_ForEmpAttendance where empno='" & sEmpID & "' and statusid=1"
                        cmdShift.Connection = CON1
                        rdShift = cmdShift.ExecuteReader
                        While rdShift.Read
                            If Not IsDBNull(rdShift(0)) Then
                                ShiftID = rdShift(0)
                                MustWork = rdShift(1)
                            Else
                                ShiftID = 0
                            End If

                        End While


                        rdShift.Close()
                        cmdShift.Dispose()


                        Dim cmdMulti As New SqlClient.SqlCommand
                        Dim rdMulti As SqlClient.SqlDataReader
                        Dim LastTime As DateTime
                        Dim TimePassed As Integer

                        cmdMulti.CommandText = "select entrydate from  tableempattendance where empno='" & sEmpID & "' and outdate is null and convert(date,entrydate)='" & Ds.ToString("MM/dd/yyyy") & "'"
                        cmdMulti.Connection = CON1

                        rdMulti = cmdMulti.ExecuteReader
                        While rdMulti.Read
                            LastTime = rdMulti(0)
                        End While
                        rdMulti.Close()
                        cmdMulti.Dispose()

                        TimePassed = DateDiff(DateInterval.Minute, LastTime, Ds)
                        If TimePassed >= 10 Then
                            WorkHours = TimePassed / 60
                            Extra = 0
                            Less = 0
                            If WorkHours > MustWork Then
                                Extra = WorkHours - MustWork
                            ElseIf WorkHours < MustWork Then
                                Less = MustWork - WorkHours
                            Else
                                Extra = 0
                                Less = 0

                            End If


                            '=================================================
                            'Handle RestDay Present as OT


                            cmdWeekEnd.CommandText = "select empno,weekendday from tableemployee where empno='" & sEmpID & "'"
                            cmdWeekEnd.Connection = CON1
                            rdWeekEnd = cmdWeekEnd.ExecuteReader

                            While rdWeekEnd.Read
                                DayNum = rdWeekEnd(1)
                            End While
                            rdWeekEnd.Close()
                            cmdWeekEnd.Dispose()


                            If DayNum = DayNum1 Then
                                RuleShort = "OT"
                            ElseIf DayNum <> DayNum1 And DayOfWeek.Sunday Then
                                RuleShort = "OT"
                            Else
                                RuleShort = "P"
                            End If



                            '==================================================


                            cmdSQL.CommandText = "update tableempattendance set outdate='" & Ds.ToString("MM/dd/yyyy  HH:mm:ss") & "',ruleshort='" & RuleShort & "',recordcomplete=1,machinesn='" & MachineName & "',empshiftid='" & ShiftID & "', workhours=" & WorkHours & ", mustwork=" & MustWork & ",extra1=" & Extra & ",less=" & Less & "  where empno='" & sEmpID & "' and outdate is null and convert(date,entrydate)='" & Ds.ToString("MM/dd/yyyy") & "'"
                            cmdSQL.Connection = CON1
                            cmdSQL.ExecuteScalar()
                            cmdSQL.Dispose()
                        Else

                            Dim cmdTodaySql As New SqlClient.SqlCommand
                            cmdTodaySql.CommandText = "insert into tablein values ('" & sEmpID & "','" & MakeDate & "','NA')"
                            cmdTodaySql.Connection = CON1
                            cmdTodaySql.ExecuteScalar()
                            cmdTodaySql.Dispose()

                        End If

                    End If

                End If
            End If


        Catch ex As Exception
            ListBox2.Items.Add(ex.Message.ToString)
        End Try


    End Sub





    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If bIsConnected = False Then
            Load1stMachine()
        End If
    End Sub

    Private Sub frmConAtt4_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Load1stMachine()
    End Sub
End Class
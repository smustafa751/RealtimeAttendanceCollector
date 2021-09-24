Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration




Module DBCON
    Public Con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("BostanSecurity.My.MySettings.USoftConnectionString").ConnectionString)
    Public CON1 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("BostanSecurity.My.MySettings.USoftConnectionString").ConnectionString)
    Public CON2 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("BostanSecurity.My.MySettings.USoftConnectionString").ConnectionString)
    Public CON3 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("BostanSecurity.My.MySettings.USoftConnectionString").ConnectionString)
    Public CON4 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("BostanSecurity.My.MySettings.USoftConnectionString").ConnectionString)
    Public CON5 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("BostanSecurity.My.MySettings.USoftConnectionString").ConnectionString)
    Public CON6 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("BostanSecurity.My.MySettings.USoftConnectionString").ConnectionString)

    'Public Con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("zString").ConnectionString)
    'Public CON1 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("zString").ConnectionString)
    'Public CON2 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("zString").ConnectionString)
    'Public CON3 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("zString").ConnectionString)
    'Public CON4 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("zString").ConnectionString)
    'Public CON5 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("zString").ConnectionString)
    'Public CON6 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("zString").ConnectionString)

    'Public Con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("volta").ConnectionString)
    'Public CON1 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("volta").ConnectionString)
    'Public CON2 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("volta").ConnectionString)
    'Public CON3 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("volta").ConnectionString)
    'Public CON4 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("volta").ConnectionString)
    'Public CON5 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("volta").ConnectionString)
    'Public CON6 As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("volta").ConnectionString)


    Public ManageEmps As Boolean

    Public EmpNo As String
    Public Gender As String
    Public JobT As String
    Public EName As String
    Public NIC As String
    Public ShiftName As String

    Public UserType As Integer
    'Public isReadOnly As Boolean
    'Public isLimited As Boolean
    Public AttRep As Boolean
    Public EmpRep As Boolean


    


    Public Sub OpenCon()
        Try
            '    Con.ConnectionString = "Integrated Security=SSPI;Data Source=home;Initial Catalog=USoft"
            Con.Open()
        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub OpenCon2()
        Try
            '   CON2.ConnectionString = "Integrated Security=SSPI;Data Source=home;Initial Catalog=USoft"
            CON2.Open()
        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub OpenCon1()
        Try
            '  CON1.ConnectionString = "Integrated Security=SSPI;Data Source=home;Initial Catalog=USoft"
            CON1.Open()
        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub OpenCon3()
        Try
            ' CON3.ConnectionString = "Integrated Security=SSPI;Data Source=home;Initial Catalog=USoft"
            CON3.Open()
        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub OpenCon4()
        Try
            ' CON3.ConnectionString = "Integrated Security=SSPI;Data Source=home;Initial Catalog=USoft"
            CON4.Open()
        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try

    End Sub


    Public Sub OpenCon5()
        Try
            ' CON3.ConnectionString = "Integrated Security=SSPI;Data Source=home;Initial Catalog=USoft"
            CON5.Open()
        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try

    End Sub


    Public Sub OpenCon6()
        Try
            ' CON3.ConnectionString = "Integrated Security=SSPI;Data Source=home;Initial Catalog=USoft"
            CON6.Open()
        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try

    End Sub
End Module

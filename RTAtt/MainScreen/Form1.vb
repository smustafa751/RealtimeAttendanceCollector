Public Class frmContainer
    Dim RunAt As String
    Dim Running As String

    Private Sub frmContainer_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Dim frmCon1 As New frmConAtt1
        'Dim frmCon2 As New frmConAtt2
        'Dim frmCon3 As New frmConAtt3
        'Dim frmCon4 As New frmConAtt4
        'Dim frmCon5 As New frmConAtt5
        'Dim frmCon6 As New frmConAtt6
        'Dim frmCon7 As New frmConAtt7
        'Dim frmCon8 As New frmConAtt8
        'Dim frmCon9 As New frmConAtt9
        'Dim frmCon10 As New frmConAtt010
        'Dim frmCon11 As New frmConAtt011
        'Dim frmCon12 As New frmConAtt012
        'Dim frmCon13 As New frmConAtt013
        'Dim frmCon14 As New frmConAtt014
        'Dim frmCon15 As New frmConAtt015

        'frmCon1.MdiParent = Me
        'frmCon2.MdiParent = Me
        'frmCon3.MdiParent = Me
        'frmCon4.MdiParent = Me
        'frmCon5.MdiParent = Me
        'frmCon6.MdiParent = Me
        'frmCon7.MdiParent = Me
        'frmCon8.MdiParent = Me
        'frmCon9.MdiParent = Me
        'frmCon10.MdiParent = Me
        'frmCon11.MdiParent = Me
        'frmCon12.MdiParent = Me
        'frmCon13.MdiParent = Me
        'frmCon14.MdiParent = Me
        'frmCon15.MdiParent = Me

        'frmCon1.Show()
        'frmCon2.Show()
        'frmCon3.Show()
        'frmCon4.Show()
        'frmCon5.Show()
        'frmCon6.Show()
        'frmCon7.Show()
        'frmCon8.Show()
        'frmCon9.Show()
        'frmCon10.Show()
        'frmCon11.Show()
        'frmCon12.Show()
        'frmCon13.Show()
        'frmCon14.Show()
        'frmCon15.Show()


        'Me.LayoutMdi(System.Windows.Forms.MdiLayout.TileHorizontal)


    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        RunAt = "01:36"
        Running = DateTime.Now.ToString("HH:mm")
        If Running = RunAt Then
            Dim frmA As New frmAbsent
            frmA.MdiParent = Me
            frmA.Show()
            Timer1.Enabled = False
            Timer2.Enabled = True
            'Me.Text = "timer1 is off"
        End If


    End Sub

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Running = DateTime.Now.ToString("HH:mm")
        If Running = "01:37" Then
            Timer1.Enabled = True
            Timer2.Enabled = False
            'Me.Text = Me.Text + "timer1 is on, timer2 is off"
        End If

    End Sub
End Class

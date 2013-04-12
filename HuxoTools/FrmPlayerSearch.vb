Public Class FrmPlayerSearch

    Private Sub CheckPlayerInfo()
        Form1.TxtPlayerLookup.Text = TxtPlayer.Text
        Form1.BgLookup.RunWorkerAsync()
        Form1.TabControl1.SelectedIndex = 0
    End Sub

    ' Controls
    Private Sub FrmPlayerSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim x As Integer = My.Computer.Screen.WorkingArea.Width - 229
        Dim y As Integer = My.Computer.Screen.WorkingArea.Height - 66

        Me.Location = New Point(x, y)
    End Sub

    Private Sub BtnLookup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLookup.Click
        CheckPlayerInfo()

        Me.Close()
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    Private Sub TxtPlayer_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPlayer.KeyUp
        If (e.KeyValue = Keys.Enter) Then
            CheckPlayerInfo()

            Me.Close()

        ElseIf (e.KeyValue = Keys.Escape) Then
            Me.Close()
        End If
    End Sub
End Class
Public Class Form1

#Region " Player Lookup "

    Dim Lookup As New LookupPlayer

    Private Sub BtnLookup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLookup.Click
        If (My.Computer.Network.IsAvailable = False) Then
            MessageBox.Show("You need to be connected to the Internet.", "Error", MessageBoxButtons.OK, _
                            MessageBoxIcon.Error)
            Exit Sub
        End If

        If (TxtPlayerLookup.Text <> Nothing) Then
            BtnLookup.Enabled = False
            ProgressBar1.Show()
            BgLookup.RunWorkerAsync()
        End If
    End Sub ' If change this...

    Private Sub TxtPlayerLookup_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPlayerLookup.KeyUp
        If (My.Computer.Network.IsAvailable = False) Then
            MessageBox.Show("You need to be connected to the Internet.", "Error", MessageBoxButtons.OK, _
                            MessageBoxIcon.Error)
            Exit Sub
        End If

        If (e.KeyValue = Keys.Enter) Then
            If (TxtPlayerLookup.Text <> Nothing) Then
                BtnLookup.Enabled = False
                ProgressBar1.Show()
                BgLookup.RunWorkerAsync()
            End If
        End If
    End Sub ' ...change this

    Private Sub LinkGuildName_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkGuildName.LinkClicked
        If (Lookup.guildLookupUrl <> Nothing) Then
            System.Diagnostics.Process.Start(Lookup.guildLookupUrl)
        End If
    End Sub

    Private Sub LinkLookupOnline_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLookupOnline.LinkClicked
        If (Lookup.PlayerLookupUrl IsNot Nothing) Then
            System.Diagnostics.Process.Start(Lookup.PlayerLookupUrl)
        End If
    End Sub

    Private Sub BgLookup_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgLookup.DoWork
        Lookup.GetPlayerInformation(TxtPlayerLookup.Text)
    End Sub

    Private Sub BgLookup_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgLookup.RunWorkerCompleted
        Lookup.ParseInformation()
        TxtPlayerLookup.Clear()
        BtnLookup.Enabled = True
        ProgressBar1.Hide()
        If (Me.Visible = False) Then
            Me.Show()
        End If
    End Sub

#End Region

#Region " VIP List "

    Dim VIP As New vip

    Private Sub VIPStatus()
        GroupPlayerControl.Enabled = False
        ProgressBar2.Show()

        LblVIPStatus.Text = "Checking..."
    End Sub

    Private Sub BtnAddPlayer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddPlayer.Click
        If (TxtAddPlayerName.Text <> Nothing) And (ComboAddWorld.Text <> Nothing) Then

            My.Computer.FileSystem.WriteAllText("VIP.txt", ComboVipImages.SelectedIndex & "," & _
                                                TxtAddPlayerName.Text & "," & ComboAddWorld.Text & vbNewLine, True)

            VIP.LoadVIPList()

            TxtAddPlayerName.Clear()
            ComboAddWorld.Text = Nothing
            ComboVipImages.SelectedItem = Nothing
        End If
    End Sub

    Private Sub BtnRemoveSel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRemoveSel.Click
        If (ListVIP.SelectedItems.Count = 0) Then
            MessageBox.Show("You must select a player to remove.", "Error", MessageBoxButtons.OK)
        Else
            VIP.RemovePlayer(ListVIP.SelectedItems.Item(0).Index)
        End If
    End Sub

    Private Sub BtnVipEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnVipEdit.Click
        If (ListVIP.SelectedItems.Count > 0) Then
            FrmEditVip.ComboVipImages.Items.Add("None")
            VIP.Edit(ListVIP.SelectedItems(0).Index)
        Else
            MessageBox.Show("Please select a player to edit.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub BtnUpdateVIP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUpdateVIP.Click
        If (ListVIP.Items.Count = 0) Then
            MessageBox.Show("You need to add players to the VIP list.", "Error", MessageBoxButtons.OK, _
                            MessageBoxIcon.Error)
            Exit Sub

        ElseIf (My.Computer.Network.IsAvailable = False) Then
            MessageBox.Show("You need to be connected to the Internet.", "Error", MessageBoxButtons.OK, _
                            MessageBoxIcon.Error)
            Exit Sub
        End If

        If (BgVip.IsBusy = False) Then
            VIPStatus()
            VIP.BeforeDownload()
            BgVip.RunWorkerAsync()
            BtnUpdateVIP.Enabled = False
        End If
    End Sub ' if change this...

    Private Sub TimerVIP_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerVIP.Tick
        If (ListVIP.Items.Count = 0) Then
            MessageBox.Show("You need to add players to the VIP list.", "Error", MessageBoxButtons.OK, _
                            MessageBoxIcon.Error)
            Exit Sub

        ElseIf (My.Computer.Network.IsAvailable = False) Then
            MessageBox.Show("You need to be connected to the Internet.", "Error", MessageBoxButtons.OK, _
                            MessageBoxIcon.Error)
            Exit Sub
        End If

        If (BgVip.IsBusy = False) Then
            VIPStatus()
            VIP.BeforeDownload()
            BtnUpdateVIP.Enabled = False
            BgVip.RunWorkerAsync()
        End If
    End Sub ' ...change this

    Private Sub ChkVIPAutoUpdate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkVIPAutoUpdate.CheckedChanged
        If (ChkVIPAutoUpdate.Checked) Then
            TimerVIP.Start()
            EnableVIPToolStripMenuItem.Checked = True
        Else
            TimerVIP.Stop()
            EnableVIPToolStripMenuItem.Checked = False
        End If
    End Sub

    Private Sub BgVip_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgVip.DoWork
        VIP.DownloadOnlinePlayersList()
    End Sub

    Private Sub BgVip_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgVip.RunWorkerCompleted
        VIP.AfterDownload()
        BtnUpdateVIP.Enabled = True
        GroupPlayerControl.Enabled = True
        ProgressBar2.Hide()

        LblVIPStatus.Text = "Last Check:" & TimeOfDay
    End Sub

#End Region

#Region " Misc "

    Dim misc As New Misc

    ' Kill Counter
    Private Sub BtnKillCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnKillCount.Click
        misc.KillCounter()
    End Sub

    Private Sub RichKillCounter_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichKillCounter.TextChanged
        If (RichKillCounter.Text = Nothing) Then
            BtnKillCount.Enabled = False
        Else
            BtnKillCount.Enabled = True
        End If
    End Sub

#End Region

#Region " Server Log Parser "

    Dim sLogParser As New ServerLogParser

    Private Sub BtnServerLogParse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnServerLogParse.Click
        sLogParser.ParseServerLog()
    End Sub

    Private Sub RichServerLog_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichServerLog.TextChanged
        If (RichServerLog.Text = Nothing) Then
            BtnServerLogParse.Enabled = False
        Else
            BtnServerLogParse.Enabled = True
        End If
    End Sub

#End Region

#Region " Notify Icon Context Menu "

    Private Sub LookupPlayerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LookupPlayerToolStripMenuItem.Click
        FrmPlayerSearch.Show()
    End Sub

    Private Sub EnableVIPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnableVIPToolStripMenuItem.Click
        If (EnableVIPToolStripMenuItem.Checked) Then
            TimerVIP.Start()
            ChkVIPAutoUpdate.Checked = True
        Else
            TimerVIP.Stop()
            ChkVIPAutoUpdate.Checked = False
        End If
    End Sub

    Private Sub ShowProgramToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowProgramToolStripMenuItem.Click
        If (ShowProgramToolStripMenuItem.Text = "Show Program") Then
            Me.Visible = True
            ShowProgramToolStripMenuItem.Text = "Hide Program"
        Else
            Me.Visible = False
            ShowProgramToolStripMenuItem.Text = "Show Program"
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

#End Region


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        VIP.LoadVIPList()

        ' Checks for already running instance
        Dim RunningProcess() As Process = System.Diagnostics.Process.GetProcessesByName(Application.ProductName)
        If (RunningProcess.Length > 1) Then
            MessageBox.Show("You can only have one instance of HuxoTools running at one time." & vbNewLine & _
                            "Application will now exit.", "Error - Program already running", MessageBoxButtons.OK, _
                            MessageBoxIcon.Error)

            Application.Exit()
        End If

        ' VIP images
        Dim images(ImageListVip.Images.Count - 1) As String
        For i As Integer = 0 To ImageListVip.Images.Count - 1
            images(i) = "Item " & i.ToString
        Next

        ComboVipImages.Items.AddRange(images)
        ComboVipImages.DrawMode = DrawMode.OwnerDrawVariable
        ComboVipImages.ItemHeight = 20
    End Sub

    Private Sub ComboVipImages_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboVipImages.DrawItem
        If (e.Index <> -1) Then
            e.Graphics.DrawImage(ImageListVip.Images(e.Index), e.Bounds.Left, e.Bounds.Top)
        End If
    End Sub

    Private Sub ComboVipImages_MeasureItem(ByVal sender As Object, ByVal e As System.Windows.Forms.MeasureItemEventArgs) Handles ComboVipImages.MeasureItem
        e.ItemHeight = ImageListVip.ImageSize.Height
        e.ItemWidth = ImageListVip.ImageSize.Width
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        NotifyIcon1.Visible = False
    End Sub

    Private Sub Form1_LocationChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.LocationChanged
        If (Me.WindowState = FormWindowState.Minimized) Then
            Me.Visible = False
        End If
    End Sub

    ' Notify Icon
    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        If (Me.Visible) Then
            Me.Visible = False
        Else
            Me.Visible = True
            Me.WindowState = FormWindowState.Normal
            'Me.TopMost = True
            'Me.TopMost = False
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://www.huxotools.webs.com")
    End Sub

    Private Sub LinkDonate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkDonate.LinkClicked
        System.Diagnostics.Process.Start("http://www.huxotools.webs.com/about.html")
    End Sub
End Class

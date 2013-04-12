Public Class vip
    Dim Players As New ArrayList
    Public Worlds As New ArrayList

    Private Sub CompileVIPList()
        Players.Clear()
        Worlds.Clear()

        ' Compiles a list of worlds and players
        Dim Item As ListViewItem
        For I As Integer = 0 To Form1.ListVIP.Items.Count - 1
            Item = Form1.ListVIP.Items.Item(I)

            Dim p As String = Item.SubItems(0).Text

            If (p.Contains(" ")) Then
                p = p.Replace(" ", "+")
            End If

            Players.Add(p)

            If (Worlds.Contains(Item.SubItems(1).Text) = False) Then
                Worlds.Add(Item.SubItems(1).Text)
            End If
        Next
    End Sub

    Public Sub DownloadOnlinePlayersList()
        If (IO.File.Exists("rawonline.txt")) Then IO.File.Delete("rawonline.txt")

        For I As Integer = 0 To Worlds.Count - 1
            Dim url As String = "http://www.tibia.com/community/?subtopic=whoisonline&world=" & Worlds.Item(I)

            If (IO.File.Exists(Worlds.Item(I) & ".txt")) Then IO.File.Delete(Worlds.Item(I) & ".txt")

            My.Computer.Network.DownloadFile(url, Worlds.Item(I) & ".txt")

            ' Starts compiling online list
            'Form1.RichRawOnline.AppendText(My.Computer.FileSystem.ReadAllText(Worlds.Item(I) & ".txt"))
            IO.File.AppendAllText("rawonline.txt", IO.File.ReadAllText(Worlds.Item(I) & ".txt"))

            IO.File.Delete(Worlds.Item(I) & ".txt")
        Next

        Worlds.Clear()
    End Sub

    Private Sub CheckForOnlinePlayers()
        Form1.RichRawOnline.LoadFile("rawonline.txt", RichTextBoxStreamType.PlainText)
        IO.File.Delete("rawonline.txt")
        For i As Integer = 0 To Players.Count - 1

            'If (Form1.RichRawOnline.Text.ToLower.Contains(Players.Item(i).ToString.ToLower)) Then
            If (Form1.RichRawOnline.Text.ToLower.Contains(Players.Item(i).ToString.ToLower)) Then

                Dim p As String = Players.Item(i)
                p = p.Replace("+", " ")

                For x As Integer = 0 To Form1.ListVIP.Items.Count - 1
                    If (Form1.ListVIP.Items.Item(x).SubItems(0).Text = p) Then
                        Dim onlineP As ListViewItem = Form1.ListVIP.Items.Item(x)
                        Form1.ListVIP.Items.Remove(onlineP)
                        onlineP.ForeColor = Color.LimeGreen
                        Form1.ListVIP.Items.Insert(0, onlineP)
                        Exit For
                    End If
                Next

            End If

        Next

        Form1.RichRawOnline.Clear()
    End Sub

    Public Sub BeforeDownload()
        LoadVIPList()
        CompileVIPList()
    End Sub

    Public Sub AfterDownload()
        CheckForOnlinePlayers()
    End Sub

    Public Sub RemovePlayer(ByVal Index As Integer)
        Dim playerName As String = Form1.ListVIP.Items.Item(Index).Text

        Dim r As New IO.StreamReader("VIP.txt")
        Dim line As String = r.ReadLine

        Dim newList As String = Nothing

        Do Until r.EndOfStream
            Dim tmp() As String = Split(line, ",")

            If (tmp(1) <> playerName) Then
                newList += tmp(0) & "," & tmp(1) & "," & tmp(2) & vbNewLine
            End If

            line = r.ReadLine
        Loop
        r.Dispose()
        r.Close()

        My.Computer.FileSystem.WriteAllText("VIP.txt", newList, False)
        LoadVIPList()
    End Sub

    Public Sub LoadVIPList()
        If (IO.File.Exists("VIP.txt") = False) Then
            Exit Sub
        End If

        Form1.ListVIP.Items.Clear()

        Dim reader As New IO.StreamReader("VIP.txt")

        Dim line As String = reader.ReadLine()

        Do Until line = Nothing
            Dim tmp() As String = Split(line, ",")
            Dim VIPAdd As New ListViewItem(tmp(1))
            VIPAdd.SubItems.Add(tmp(2))
            VIPAdd.ImageIndex = tmp(0)
            Form1.ListVIP.Items.Add(VIPAdd)

            line = reader.ReadLine()
        Loop

        reader.Dispose()
        reader.Close()
    End Sub

    Public Sub Edit(ByVal Index As Integer)
        Dim playerName As String = Form1.ListVIP.Items.Item(Index).Text

        Dim r As New IO.StreamReader("VIP.txt")
        Dim line As String = r.ReadLine

        Do Until r.EndOfStream
            Dim tmp() As String = Split(line, ",")

            If (tmp(1) = playerName) Then
                'FrmEditVip.ComboVipImages.SelectedIndex = tmp(0)
                FrmEditVip.TxtName.Text = tmp(1)
                FrmEditVip.ComboWorld.Text = tmp(2)

                FrmEditVip.TextBox1.Text = line
                r.Dispose()
                r.Close()
                FrmEditVip.ShowDialog()
                LoadVIPList()
                Exit Do
            End If
            line = r.ReadLine
        Loop

        r.Dispose()
        r.Close()
    End Sub
End Class

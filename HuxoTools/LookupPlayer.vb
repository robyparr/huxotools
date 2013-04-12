Public Class LookupPlayer

    Public guildLookupUrl As String = Nothing
    Public PlayerLookupUrl As String = Nothing

    Public Sub GetPlayerInformation(ByVal Player As String)
        Try
            ' Replaces spaces with + signs
            Player = Player.Replace(" ", "+")

            ' URL to download from
            PlayerLookupUrl = "http://www.tibia.com/community/?subtopic=characters&name=" + Player

            ' Downloads information
            My.Computer.Network.DownloadFile(PlayerLookupUrl, "char.txt", "", "", False, 100000, True)
        Catch ex As Exception
            MessageBox.Show(ex.ToString & vbNewLine & vbNewLine & "Please report this error to the programmer.", "error")
        End Try
    End Sub

    Public Sub ParseInformation()
        Try
            '' Reads raw text
            Form1.RichRawText.LoadFile("char.txt", RichTextBoxStreamType.PlainText)

            ' Deletes file
            If (IO.File.Exists("char.txt")) Then IO.File.Delete("char.txt")

            ' Clears fields
            Form1.TxtName.Clear()
            Form1.TxtSex.Clear()
            Form1.TxtTown.Clear()
            Form1.TxtVoc.Clear()
            Form1.TxtWorld.Clear()
            Form1.TxtLevel.Clear()
            Form1.TxtGuild.Clear()
            Form1.TxtLastLogin.Clear()
            Form1.TxtAccountStatus.Clear()
            Form1.LinkGuildName.Text = "GuildName"
            Form1.LinkGuildName.Visible = False
            Form1.LblGuild.Visible = False
            guildLookupUrl = Nothing

            ' Starts search
            Dim StartPosition As Integer
            Dim SearchType As CompareMethod = CompareMethod.Text

            ' Name
            StartPosition = InStr(1, Form1.RichRawText.Text, "Name:</TD><TD>", SearchType)
            If ((StartPosition <> 0)) Then
                If (Form1.TxtPlayerLookup.Text.Length = 0) Then
                    Form1.RichRawText.Select(StartPosition + 13, FrmPlayerSearch.TxtPlayer.Text.Length)
                Else
                    Form1.RichRawText.Select(StartPosition + 13, Form1.TxtPlayerLookup.Text.Length)
                End If
                If (Form1.RichRawText.SelectedText.Contains("<")) Then
                    MessageBox.Show("No player with that name has been found.", "HuxoTools - Player Lookup", _
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                Else
                    Form1.TxtName.Text = Form1.RichRawText.SelectedText
                End If
            End If

            ' Sex
            StartPosition = InStr(1, Form1.RichRawText.Text, "Sex:</TD><TD>", SearchType)
            If ((StartPosition <> 0)) Then
                For i As Integer = 2 To 100
                    Form1.RichRawText.Select(StartPosition + 12, i)
                    If (Form1.RichRawText.SelectedText.Contains("<")) Then
                        Form1.RichRawText.Select(StartPosition + 12, i - 1)
                        Exit For
                    End If
                Next
                Form1.TxtSex.Text = Form1.RichRawText.SelectedText
            End If

            ' Vocation
            StartPosition = InStr(1, Form1.RichRawText.Text, "Profession:</TD><TD>", SearchType)
            If ((StartPosition <> 0)) Then
                For i As Integer = 2 To 100
                    Form1.RichRawText.Select(StartPosition + 19, i)
                    If (Form1.RichRawText.SelectedText.Contains("<")) Then
                        Form1.RichRawText.Select(StartPosition + 19, i - 1)
                        Exit For
                    End If
                Next
                Form1.TxtVoc.Text = Form1.RichRawText.SelectedText
            End If

            ' Level
            StartPosition = InStr(1, Form1.RichRawText.Text, "Level:</TD><TD>", SearchType)
            If ((StartPosition <> 0)) Then
                For i As Integer = 2 To 100
                    Form1.RichRawText.Select(StartPosition + 14, i)
                    If (Form1.RichRawText.SelectedText.Contains("<")) Then
                        Form1.RichRawText.Select(StartPosition + 14, i - 1)
                        Exit For
                    End If
                Next
                Form1.TxtLevel.Text = Form1.RichRawText.SelectedText
            End If

            ' World
            StartPosition = InStr(1, Form1.RichRawText.Text, "World:</TD><TD>", SearchType)
            If ((StartPosition <> 0)) Then
                For i As Integer = 2 To 100
                    Form1.RichRawText.Select(StartPosition + 14, i)
                    If (Form1.RichRawText.SelectedText.Contains("<")) Then
                        Form1.RichRawText.Select(StartPosition + 14, i - 1)
                        Exit For
                    End If
                Next
                Form1.TxtWorld.Text = Form1.RichRawText.SelectedText
            End If

            ' Residence
            StartPosition = InStr(1, Form1.RichRawText.Text, "Residence:</TD><TD>", SearchType)
            If ((StartPosition <> 0)) Then
                For i As Integer = 2 To 100
                    Form1.RichRawText.Select(StartPosition + 18, i)
                    If (Form1.RichRawText.SelectedText.Contains("<")) Then
                        Form1.RichRawText.Select(StartPosition + 18, i - 1)
                        Exit For
                    End If
                Next
                Form1.TxtTown.Text = Form1.RichRawText.SelectedText
            End If

            ' Comment
            'StartPosition = InStr(1, Form1.RichRawText.Text, "Comment:</TD><TD>", SearchType)
            'If (StartPosition <> 0) Then
            '    For i As Integer = 2 To 500
            '        Form1.RichRawText.Select(StartPosition + 16, i)
            '        Form1.RichRawText.SelectedText = Form1.RichRawText.SelectedText.Replace("<br />", vbNewLine)
            '        If (Form1.RichRawText.SelectedText.Contains("<")) Then
            '            Form1.RichRawText.Select(StartPosition + 16, i - 1)
            '            Form1.RichRawText.ScrollToCaret()
            '            Exit For
            '        End If
            '    Next
            '    RichComment.Text = Form1.RichRawText.SelectedText
            'End If

            ' Guild membership
            StartPosition = InStr(1, Form1.RichRawText.Text, "Guild&#160;membership:</TD><TD>", SearchType)
            If (StartPosition <> 0) Then
                For i As Integer = 2 To 100
                    Form1.RichRawText.Select(StartPosition + 30, i)
                    If (Form1.RichRawText.SelectedText.Contains(" of the ")) Then
                        Form1.RichRawText.Select(StartPosition + 30, i - 8)
                        Exit For
                    End If
                Next
                Form1.TxtGuild.Text = Form1.RichRawText.SelectedText
            End If


            ' Guild name
            StartPosition = InStr(1, Form1.RichRawText.Text, "http://www.tibia.com/community/?subtopic=guilds&page=view&GuildName=", SearchType)
            If (StartPosition <> 0) Then
                For i As Integer = 2 To 100
                    Form1.RichRawText.Select(StartPosition + 67, i)
                    If (Form1.RichRawText.SelectedText.Contains(">")) Then
                        Form1.RichRawText.Select(StartPosition + 67, i - 2)
                        Dim gn As String = Form1.RichRawText.SelectedText
                        guildLookupUrl = "http://www.tibia.com/community/?subtopic=guilds&page=view&GuildName=" & gn
                        Form1.LinkGuildName.Text = gn.Replace("+", " ")
                        Form1.LblGuild.Visible = True
                        Form1.LinkGuildName.Visible = True
                        Exit For
                    End If
                Next
            End If

            ' Last login
            StartPosition = InStr(1, Form1.RichRawText.Text, "Last login:</TD><TD>", SearchType)
            If (StartPosition <> 0) Then
                For i As Integer = 2 To 100
                    Form1.RichRawText.Select(StartPosition + 19, i)
                    If (Form1.RichRawText.SelectedText.Contains("<")) Then
                        Form1.RichRawText.Select(StartPosition + 19, i - 1)
                        Exit For
                    End If
                Next
                Form1.TxtLastLogin.Text = Form1.RichRawText.SelectedText.Replace("&#160;", " ")
            End If

            ' Account status
            StartPosition = InStr(1, Form1.RichRawText.Text, "Account&#160;Status:</TD><TD>", SearchType)
            If (StartPosition <> 0) Then
                For i As Integer = 2 To 100
                    Form1.RichRawText.Select(StartPosition + 28, i)
                    If (Form1.RichRawText.SelectedText.Contains("<")) Then
                        Form1.RichRawText.Select(StartPosition + 28, i - 1)
                        Exit For
                    End If
                Next
                Form1.TxtAccountStatus.Text = Form1.RichRawText.SelectedText
            End If

            ' Death list
            Form1.ListDeathList.Items.Clear()

            If (Form1.RichRawText.Text.Contains("Character Deaths") = False) Then Exit Sub

            Dim toDrop() As String = {"<b>", "<td>", "</b>", "</td>", "<tr>", "</tr>", _
            "<tr bgcolor=""#F1E0C6"" ><td width=""25%"" valign=""top"" >", "&#160;", _
            "<a href=""http://www.tibia.com/community/?subtopic=characters&name=", " >", "</a>", _
            "<tr bgcolor=""#D4C0A1"" <td width=""25%"" valign=""top""", "<br/>", "</table>", _
            "<td width=""25%"" valign=""top""", "  "}

            StartPosition = InStr(1, Form1.RichRawText.Text, "<b>Character Deaths</b></td>", SearchType)
            Dim str As String = Nothing
            Dim Counter As Integer = 1

            Do Until (Form1.RichRawText.SelectedText.Contains("</td></tr></table><br/><br/>"))
                Form1.RichRawText.Select(StartPosition + 32, Counter)

                If (Form1.RichRawText.SelectedText.Contains("</td></tr></table><br/><br/>")) Then
                    Form1.RichRawText.Select(StartPosition + 32, Counter - 28)
                    str = Form1.RichRawText.SelectedText
                End If

                Counter += 1
            Loop

            For Each Item As String In toDrop
                If (str.Contains(Item)) Then
                    str = str.Replace(Item, " ")
                End If
            Next

            ' Mar 11 2010, 22:09:49 CET Slain at Level 227 by Bernis" Bernis , Decoy+Octopus" Decoy Octopus , 
            ' Milivin" Milivin , Scott+yellow" Scott yellow and Sobek+pogromca" Sobek pogromca .   

            While (str.Contains(""""))
                StartPosition = InStr(1, str, """", SearchType)
                'Dim last As Integer = StartPosition
                If (StartPosition <> 0) Then
                    For i As Integer = 5 To 100
                        Dim str2 As String = str.Substring(StartPosition - 1, i)
                        If (str2.EndsWith(" ,")) Or (str2.EndsWith(" .")) Or (str2.EndsWith(" and")) Then
                            If (str2.EndsWith(" and")) Then
                                str = str.Remove(StartPosition - 1, i - 4)
                            Else
                                str = str.Remove(StartPosition - 1, i - 1)
                            End If
                            Exit For
                        End If
                    Next
                End If
            End While
            str = str.Replace("+", " ")
            'If (str.Contains(""" ")) Then
            '    Dim start As Integer = str.IndexOf("by ") + 3
            '    Dim finish As Integer = InStr(str, """ ", CompareMethod.Text)
            '    Dim Remove As String = str.Substring(start, (finish - start))
            '    str = str.Replace(Remove, "")

            '    Dim start2 As Integer = str.IndexOf("and ") + 4
            '    Dim finish2 As Integer = InStr(str, """ ", CompareMethod.Text)
            '    Dim Remove2 As String = str.Substring(start2, (finish2 - start2))
            '    str = str.Replace(Remove2, "")
            'End If

            Dim tmp() As String = Split(str, ".")

            Dim bgC As Integer = 0
            For z As Integer = 0 To tmp.Length - 2
                tmp(z) = tmp(z).Trim(" ")
                Dim r As New ListViewItem(tmp(z).Substring(0, 25))
                r.SubItems.Add(tmp(z).Substring(26))

                If (bgC = 0) Then
                    r.BackColor = Color.LightGoldenrodYellow
                    bgC = 1
                Else
                    r.BackColor = Color.White
                    bgC = 0
                End If
                Form1.ListDeathList.Items.Add(r)
            Next
            ' <b>Character Deaths</b></td></tr><tr bgcolor="#F1E0C6" ><td width="25%" 
            ' valign="top" >Mar&#160;14&#160;2010,&#160;17:58:30&#160;CET</td><td>Killed  at Level 95 by 
            ' <a href="http://www.tibia.com/community/?subtopic=characters&name=Asredynlas" >Asredynlas</a> 
            ' and <a href="http://www.tibia.com/community/?subtopic=characters&name=Meradilver+Alirax" >Meradilver
            ' &#160;Alirax</a>.</td></tr><tr bgcolor="#D4C0A1" ><td width="25%" valign="top" >Mar&#160;08&#160;2010,
            ' &#160;01:47:45&#160;CET</td><td>Died at Level 95 by a grim reaper.</td></tr></table><br/><br/>
            ' <TABLE BORDER=0 CELLSPACING=1 CELLPADDING=4 WIDTH=100%><TR BGCOLOR=#505050><TD COLSPAN=2 CLASS=white>
            ' <B>Account Information</B></TD></TR><TR BGCOLOR=#F1E0C6><TD WIDTH=20%>Created:</TD><TD>Nov&#160;18&#160;
            ' 2005,&#160;02:01:21&#160;CET</TD></TR></TABLE>
        Catch ex As Exception
            MessageBox.Show(ex.ToString & vbNewLine & vbNewLine & "Please report this to the programmer.", _
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class

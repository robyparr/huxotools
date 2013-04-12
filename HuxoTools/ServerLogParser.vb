Public Class ServerLogParser

    Dim timeStamps As New ArrayList
    Dim damageList As New ArrayList
    Dim dtItems As New Dictionary(Of String, Integer)
    Dim dtCreatures As New Dictionary(Of String, Integer)
    '00:55 Loot of a goblin: 3 gold coins, a plate armor

    Public Sub ParseServerLog()
        Try
            ' Clears lists
            Form1.ListCreaturesKilled.Items.Clear()
            Form1.ListLoot.Items.Clear()

            ' Set up temp file for loot parsing
            Form1.RichServerLog.SaveFile("lootTmp", RichTextBoxStreamType.PlainText)
            IO.File.AppendAllText("lootTmp", vbNewLine & "--End of Server Log--")

            Dim r As New IO.StreamReader("lootTmp")
            Dim line As String = r.ReadLine

            Dim params() As String = {"deal", "damage", "Loot"}
            Do Until r.EndOfStream
                For x As Integer = 0 To params.Length - 1
                    If (line.Contains(params(x))) Then
                        Exit For

                    ElseIf (x = params.Length - 1) Then
                        line = r.ReadLine
                        Continue Do
                    End If
                Next
                'If (line = Nothing) Then
                '    line = r.ReadLine
                '    Continue Do
                'Else
                '    For x As Integer = 0 To 9
                '        If (line.StartsWith(x)) Then
                '            Exit For

                '        ElseIf (line.StartsWith(x) = False) And (x = 9) Then
                '            line = r.ReadLine
                '            Continue Do
                '        End If
                '    Next
                'End If

                ' Extracts timestamp
                Dim timeStamp As String = line.Substring(0, line.IndexOf(" "))
                timeStamps.Add(timeStamp)

                ' Checks to see the type of message being parsed
                If (line.Contains("deal")) And (line.Contains("damage")) Then ' Damage
                    damageList.Add(line.Substring(15, line.IndexOf(" da") - 15))

                    line = r.ReadLine
                    Continue Do

                ElseIf (line.Contains("Loot")) Then ' Loot
                    ParseLoot(line)
                End If


                line = r.ReadLine
            Loop

            ' Displays loot results
            Dim bgC As Integer = 0
            For Each kvp As KeyValuePair(Of String, Integer) In dtItems
                Dim lootItem As New ListViewItem(kvp.Key)
                lootItem.SubItems.Add(kvp.Value)

                If (bgC = 0) Then
                    lootItem.BackColor = Color.LightGoldenrodYellow
                    bgC = 1
                Else
                    lootItem.BackColor = Color.White
                    bgC = 0
                End If

                If (kvp.Key = "gold coins") Then
                    lootItem.ForeColor = Color.LimeGreen
                ElseIf (kvp.Key = "nothing") Then
                    lootItem.ForeColor = Color.Red
                End If

                Form1.ListLoot.Items.Add(lootItem)
            Next

            ' Displays creature results
            bgC = 0
            For Each creature As KeyValuePair(Of String, Integer) In dtCreatures
                Dim creatureKilledItem As New ListViewItem(creature.Key)
                creatureKilledItem.SubItems.Add(creature.Value)

                If (bgC = 0) Then
                    creatureKilledItem.BackColor = Color.LightGoldenrodYellow
                    bgC = 1
                Else
                    creatureKilledItem.BackColor = Color.White
                    bgC = 0
                End If

                Form1.ListCreaturesKilled.Items.Add(creatureKilledItem)
            Next

            ' Display damage results
            If (Form1.RichServerLog.Text.Contains("damage")) Then
                ParseDamage()
            End If

            ' Cleanup
            r.Dispose()
            r.Close()
            If (IO.File.Exists("lootTmp")) Then IO.File.Delete("lootTmp")
            dtCreatures.Clear()
            dtItems.Clear()

            ' Runs other processes
            If (timeStamps.Count > 0) Then
                GetHuntingTimes()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error - Report to Programmer", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub GetHuntingTimes()
        ' Gets total time spent hunting
        Dim hStart As DateTime = Convert.ToDateTime(timeStamps.Item(0))
        Dim hEnd As DateTime = Convert.ToDateTime(timeStamps.Item(timeStamps.Count - 1))
        Dim hLength As TimeSpan = hStart - hEnd

        Form1.LblHuntStart.Text = "Start:    " & hStart.ToLongTimeString
        Form1.LblHuntEnd.Text = "End:     " & hEnd.ToLongTimeString
        Form1.lblHuntLength.Text = "Length: " & Strings.Left(hLength.Duration.ToString, 5)

        timeStamps.Clear()
    End Sub

    Private Sub ParseDamage()
        ' 22:54 You deal 245 damage to a bandit.

        ' Gets damage statistics
        Dim aDamageList() As Integer = Nothing
        Dim smallestNum As Integer = 10000
        Dim largestNum As Integer = 0
        Dim totalDmg As Integer = 0
        Dim avgDmg As Integer = 0

        For Each Item In damageList
            If (Item > largestNum) Then
                largestNum = Item
            End If

            If (Item < smallestNum) Then
                smallestNum = Item
            End If

            totalDmg += Item
        Next

        avgDmg = Math.Round(totalDmg / damageList.Count, 0)

        Form1.LblDamageMin.Text = "Min:         " & smallestNum
        Form1.LblDamageMax.Text = "Max:        " & largestNum
        Form1.LblDamageAverage.Text = "Average: " & avgDmg
    End Sub

    Private Sub ParseLoot(ByVal ServerConsoleLine As String)
        ' Gets everything after the timestamp
        Dim message As String = ServerConsoleLine.Substring(6)

        ' Grabs creature name
        If (message.Contains("Loot of an")) Then
            ParseCreature(message.Substring(message.IndexOf("of an ") + 6, message.IndexOf(":") - 11))
        Else
            ParseCreature(message.Substring(message.IndexOf("of a ") + 5, message.IndexOf(":") - 10))
        End If

        ' Gets items; everything after the first colon
        Dim loot As String = message.Substring(message.IndexOf(":") + 1)

        ' Split the items based on commas
        Dim lootList() As String = loot.Split(",")

        ' These will be removed from the front of strings, not including spaces
        Dim stringsToRemove() As String = {"an", "a", "some"}

        'dtItems = New Dictionary(Of String, Integer)
        For Each sLoot As String In lootList
            ' Adds a 's' on the end of gold coin
            If (sLoot.Contains("a gold coin")) Then sLoot += "s"
            Dim sClean As String = sLoot.Trim()
            For Each sToRemove As String In stringsToRemove
                If (sClean.StartsWith(sToRemove)) Then
                    sClean = sClean.TrimStart(sToRemove.ToCharArray).Trim()
                    Exit For
                End If
            Next
            Dim iCount As Integer = 0
            If (sClean.Contains(" ")) Then 'either two words in item name or multiple items'
                'For 3 gold coins Substring will return 3'
                'For plate armor Substring will return plate'
                If (Integer.TryParse(sClean.Substring(0, sClean.IndexOf(" ")), iCount)) Then
                    Dim sItemKey As String = sClean.Substring(iCount.ToString.Length + 1) 'adding one to iCounts length removes the space'
                    If Not dtItems.ContainsKey(sItemKey) Then
                        dtItems.Add(sItemKey, iCount)
                    Else
                        dtItems(sItemKey) += iCount
                    End If
                Else 'single multi-word item'
                    If Not dtItems.ContainsKey(sClean) Then
                        dtItems.Add(sClean, 1)
                    Else
                        dtItems(sClean) += 1
                    End If
                End If
            Else 'single item'
                If Not dtItems.ContainsKey(sClean) Then
                    dtItems.Add(sClean, 1)
                Else
                    dtItems(sClean) += 1
                End If
            End If
        Next

    End Sub

    Private Sub ParseCreature(ByVal Creature As String)
        If (dtCreatures.ContainsKey(Creature)) Then
            dtCreatures(Creature) += 1
        Else
            dtCreatures.Add(Creature, 1)
        End If
    End Sub
End Class

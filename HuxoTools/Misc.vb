Public Class Misc

    Public Sub KillCounter()
        Try
            Dim dates As New ArrayList

            Form1.RichKillCounter.SaveFile("killcounter.txt", RichTextBoxStreamType.PlainText)
            IO.File.AppendAllText("killcounter.txt", vbNewLine)

            Dim r As New IO.StreamReader("killcounter.txt")
            Dim line As String = r.ReadLine()

            Do Until (line = Nothing)
                If (Strings.Right(line, 2) <> "ok") Then
                    dates.Add(Strings.Left(line, 11))
                End If
                line = r.ReadLine()
            Loop

            r.Dispose()
            r.Close()
            If (IO.File.Exists("killcounter.txt")) Then IO.File.Delete("killcounter.txt")

            Dim today As Integer = 0
            Dim week As Integer = 0
            Dim month As Integer = 0

            Dim diff As Integer
            For Each Item In dates
                diff = DateDiff(DateInterval.Day, Item, Date.Today)
                If (diff = 0) Then
                    today += 1
                    week += 1
                    month += 1

                ElseIf (diff < 7) And (diff > 0) Then
                    week += 1
                    month += 1

                ElseIf (diff < 30) And (diff > 6) Then
                    month += 1
                End If
            Next

            Dim dayLeft As Integer = 3
            Dim weekLeft As Integer = 5
            Dim monthLeft As Integer = 10

            dayLeft -= today
            weekLeft -= week
            monthLeft -= month


            Form1.LblKillCountDay.Text = today & " (" & dayLeft & ")"
            Form1.LblKillCountWeek.Text = week & " (" & weekLeft & ")"
            Form1.LblKillCountMonth.Text = month & " (" & monthLeft & ")"

        Catch ex As Exception
            MessageBox.Show(ex.ToString & vbNewLine & vbNewLine & "Please report this to the programmer.", "Error", _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class

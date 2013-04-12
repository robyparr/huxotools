Public Class FrmEditVip

    Private Sub FrmEditVip_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' VIP images
        ComboVipImages.Items.RemoveAt(0)
        Dim images(Form1.ImageListVip.Images.Count - 1) As String
        For i As Integer = 0 To Form1.ImageListVip.Images.Count - 1
            images(i) = "Item " & i.ToString
        Next

        ComboVipImages.Items.AddRange(images)
        ComboVipImages.DrawMode = DrawMode.OwnerDrawVariable
        ComboVipImages.ItemHeight = 20
        ComboVipImages.MaxDropDownItems = Form1.ImageListVip.Images.Count
    End Sub

    Private Sub ComboVipImages_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboVipImages.DrawItem
        If (e.Index <> -1) Then
            e.Graphics.DrawImage(Form1.ImageListVip.Images(e.Index), e.Bounds.Left, e.Bounds.Top)
        End If
    End Sub

    Private Sub ComboVipImages_MeasureItem(ByVal sender As Object, ByVal e As System.Windows.Forms.MeasureItemEventArgs) Handles ComboVipImages.MeasureItem
        e.ItemHeight = Form1.ImageListVip.ImageSize.Height
        e.ItemWidth = Form1.ImageListVip.ImageSize.Width
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        TxtName.Clear()
        ComboWorld.Text = Nothing
        ComboVipImages.SelectedIndex = -1
        ComboVipImages.Items.Clear()
        Me.Close()
    End Sub

    Private Sub BtnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUpdate.Click
        Dim list As String = My.Computer.FileSystem.ReadAllText("VIP.txt")

        Dim edited As String = ComboVipImages.SelectedIndex & "," & TxtName.Text & "," & ComboWorld.Text
        list = list.Replace(TextBox1.Text, edited)

        My.Computer.FileSystem.WriteAllText("VIP.txt", list, False)

        TxtName.Clear()
        ComboWorld.Text = Nothing
        ComboVipImages.SelectedIndex = -1
        ComboVipImages.Items.Clear()

        Me.Close()
    End Sub
End Class
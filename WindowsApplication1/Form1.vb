Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim first, second As String
        Dim a, count As Integer
        count = 0
        TextBox2.Text = ""

        For a = 1 To Len(TextBox1.Text)
            first = Mid(TextBox1.Text, a, 1)
            second = Mid(TextBox1.Text, a + 6, 1)
            '检测字符串是否符合key格式
            If first = "-" And second = "-" Then
                If a - 6 >= 0 Then
                    '记录key数量
                    count = count + 1
                    '选中key
                    TextBox1.SelectionStart = a - 6
                    TextBox1.SelectionLength = 17
                    'key之间用半角逗号隔开
                    TextBox2.Text = TextBox2.Text + TextBox1.SelectedText + ","
                End If
            End If
        Next

        If count = 0 Then
            MsgBox("未找到 Key", vbCritical, "Error")
        Else
            TextBox2.Text = "!redeem " + TextBox2.Text  '转换成ASF格式
            Label1.Text = "已找到 " & count & " 个 Key"
            TextBox2.Text = TextBox2.Text.Substring(0, TextBox2.TextLength - 1)  '删除多余的逗号
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.Text <> "" Then
            Clipboard.Clear()
            Clipboard.SetText(TextBox2.Text)
            MsgBox("已复制到剪贴板", vbOKOnly, "Success")
        Else
            MsgBox("请先点击生成", vbCritical, "Error")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
    End Sub
End Class
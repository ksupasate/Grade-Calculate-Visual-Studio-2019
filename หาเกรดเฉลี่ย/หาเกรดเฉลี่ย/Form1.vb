Public Class Form1
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim Midtermscore As String = Val(TextBox1.Text)
        If (IsNumeric(Midtermscore)) And (Val(Midtermscore) > 20) Then
            MsgBox("คะแนนกลางภาคต้องไม่เกิน 20 !")
        ElseIf (IsNumeric(Midtermscore)) And (Val(Midtermscore) < 0) Then
            MsgBox("คะแนนกลางภาคต้องไม่ติดลบ ! ")
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim Finalscore As String = Val(TextBox2.Text)
        If (IsNumeric(Finalscore)) And (Val(Finalscore) > 20) Then
            MsgBox("คะแนนปลายภาคต้องไม่เกิน 20 !")
        ElseIf (IsNumeric(Finalscore)) And (Val(Finalscore) < 0) Then
            MsgBox("คะแนนปลายภาคต้องไม่ติดลบ ! ")
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim score As String = Val(TextBox3.Text)
        If (IsNumeric(score)) And (Val(score) > 60) Then
            MsgBox("คะแนนเก็บต้องไม่เกิน 60 !")
        ElseIf (IsNumeric(score)) And (Val(score) < 0) Then
            MsgBox("คะแนนเก็บต้องไม่ติดลบ ! ")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim s1, s2, s3, total As Single
        s1 = TextBox1.Text
        s2 = TextBox2.Text
        s3 = TextBox3.Text
        total = s1 + s2 + s3
        TextBox4.Text = total

        If total >= 80 And total <= 100 Then
            TextBox5.Text = "Grade 4"
        ElseIf total >= 75 And total <= 79 Then
            TextBox5.Text = "Grade 3.5"
        ElseIf total >= 70 And total <= 74 Then
            TextBox5.Text = "Grade 3"
        ElseIf total >= 65 And total <= 69 Then
            TextBox5.Text = "Grade 2.5"
        ElseIf total >= 60 And total <= 64 Then
            TextBox5.Text = "Grade 2"
        ElseIf total >= 55 And total <= 59 Then
            TextBox5.Text = "Grade 1.5"
        ElseIf total >= 50 And total <= 54 Then
            TextBox5.Text = "Grade 1"
        ElseIf total >= 1 And total <= 49 Then
            TextBox5.Text = "You Fail !"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class

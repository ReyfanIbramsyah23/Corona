Public Class Form1
    Dim sqlnya As String
    Sub nyalakancheckbox()
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        CheckBox6.Enabled = True
        CheckBox7.Enabled = True
        CheckBox8.Enabled = True
        CheckBox9.Enabled = True
        CheckBox10.Enabled = True
        CheckBox11.Enabled = True
        CheckBox12.Enabled = True
        CheckBox13.Enabled = True
        CheckBox14.Enabled = True
        CheckBox15.Enabled = True
        CheckBox16.Enabled = True
        CheckBox17.Enabled = True
        CheckBox18.Enabled = True
        CheckBox19.Enabled = True
        CheckBox20.Enabled = True
        CheckBox21.Enabled = True
        CheckBox22.Enabled = True
        CheckBox23.Enabled = True
        CheckBox24.Enabled = True
        CheckBox25.Enabled = True
        CheckBox26.Enabled = True
        CheckBox27.Enabled = True
        CheckBox28.Enabled = True
        CheckBox29.Enabled = True
        CheckBox30.Enabled = True
        CheckBox31.Enabled = True
        CheckBox32.Enabled = True
        CheckBox33.Enabled = True
        CheckBox34.Enabled = True
        CheckBox35.Enabled = True
        CheckBox36.Enabled = True
        CheckBox36.Enabled = True
        CheckBox37.Enabled = True
        CheckBox38.Enabled = True
        CheckBox39.Enabled = True
        CheckBox40.Enabled = True
        CheckBox41.Enabled = True
        CheckBox42.Enabled = True
    End Sub
    Sub matikancheckbox()
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        CheckBox5.Enabled = False
        CheckBox6.Enabled = False
        CheckBox7.Enabled = False
        CheckBox8.Enabled = False
        CheckBox9.Enabled = False
        CheckBox10.Enabled = False
        CheckBox11.Enabled = False
        CheckBox12.Enabled = False
        CheckBox13.Enabled = False
        CheckBox14.Enabled = False
        CheckBox15.Enabled = False
        CheckBox16.Enabled = False
        CheckBox17.Enabled = False
        CheckBox18.Enabled = False
        CheckBox19.Enabled = False
        CheckBox20.Enabled = False
        CheckBox21.Enabled = False
        CheckBox22.Enabled = False
        CheckBox23.Enabled = False
        CheckBox24.Enabled = False
        CheckBox25.Enabled = False
        CheckBox26.Enabled = False
        CheckBox27.Enabled = False
        CheckBox28.Enabled = False
        CheckBox29.Enabled = False
        CheckBox30.Enabled = False
        CheckBox31.Enabled = False
        CheckBox32.Enabled = False
        CheckBox33.Enabled = False
        CheckBox34.Enabled = False
        CheckBox35.Enabled = False
        CheckBox36.Enabled = False
        CheckBox36.Enabled = False
        CheckBox37.Enabled = False
        CheckBox38.Enabled = False
        CheckBox39.Enabled = False
        CheckBox40.Enabled = False
        CheckBox41.Enabled = False
        CheckBox42.Enabled = False
    End Sub

    Sub kosongkanform()
        Txtlogin.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        TextBox5.Text = ""
        Txtlogin.Focus()


    End Sub

    Sub jalan()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konekDB()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnya
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
    End Sub

    Sub matikanform()
        Txtlogin.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        ComboBox1.Enabled = False
        TextBox5.Enabled = False

    End Sub
    Sub hidupkanform()
        Txtlogin.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        ComboBox1.Enabled = True
    End Sub
    Sub tampilkandata()
        Call konekDB()
        DA = New OleDb.OleDbDataAdapter("select * from corona", CONN)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Dim jawaban1, jawaban2, jawaban3, jawaban4, jawaban5, jawaban6, jawaban7,
        jawaban8, jawaban9, jawaban10, jawaban11, jawaban12, jawaban13, jawaban14,
        jawaban15, jawaban16, jawaban17, jawaban18, jawaban19, jawaban20, jawaban21, Totaljawabanya As Integer
    Private Sub Button2_Click(ByVal sender As Object, e As EventArgs) Handles Button2.Click

        If CheckBox1.Checked = True Then
            jawaban1 = 1
        Else
            jawaban1 = 0
        End If
        If CheckBox2.Checked = True Then
            jawaban2 = 1
        Else
            jawaban2 = 0
        End If
        If CheckBox3.Checked = True Then
            jawaban3 = 1
        Else
            jawaban3 = 0
        End If
        If CheckBox4.Checked = True Then
            jawaban4 = 1
        Else
            jawaban4 = 0
        End If
        If CheckBox5.Checked = True Then
            jawaban5 = 1
        Else
            jawaban5 = 0
        End If
        If CheckBox6.Checked = True Then
            jawaban6 = 1
        Else
            jawaban6 = 0
        End If
        If CheckBox7.Checked = True Then
            jawaban7 = 1
        Else
            jawaban7 = 0
        End If
        If CheckBox8.Checked = True Then
            jawaban8 = 1
        Else
            jawaban8 = 0
        End If
        If CheckBox9.Checked = True Then
            jawaban9 = 1
        Else
            jawaban9 = 0
        End If
        If CheckBox10.Checked = True Then
            jawaban10 = 1
        Else
            jawaban10 = 0
        End If
        If CheckBox21.Checked = True Then
            jawaban11 = 1
        Else
            jawaban11 = 0
        End If
        If CheckBox22.Checked = True Then
            jawaban12 = 1
        Else
            jawaban12 = 0
        End If
        If CheckBox23.Checked = True Then
            jawaban13 = 1
        Else
            jawaban13 = 0
        End If
        If CheckBox24.Checked = True Then
            jawaban14 = 1
        Else
            jawaban14 = 0
        End If
        If CheckBox25.Checked = True Then
            jawaban15 = 1
        Else
            jawaban15 = 0
        End If
        If CheckBox26.Checked = True Then
            jawaban16 = 1
        Else
            jawaban16 = 0
        End If
        If CheckBox33.Checked = True Then
            jawaban17 = 1
        Else
            jawaban17 = 0
        End If
        If CheckBox34.Checked = True Then
            jawaban18 = 1
        Else
            jawaban18 = 0
        End If
        If CheckBox35.Checked = True Then
            jawaban19 = 1
        Else
            jawaban19 = 0
        End If
        If CheckBox36.Checked = True Then
            jawaban20 = 1
        Else
            jawaban20 = 0
        End If
        If CheckBox37.Checked = True Then
            jawaban21 = 1
        Else
            jawaban21 = 0
        End If

        Totaljawabanya = jawaban1 + jawaban2 + jawaban3 + jawaban4 + jawaban5 + jawaban6 + jawaban7 + jawaban8 + jawaban9 + jawaban10 + jawaban11 + jawaban12 +
            jawaban13 + jawaban14 + jawaban15 + jawaban16 + jawaban17 + jawaban18 + jawaban19 + jawaban20 + jawaban21
        TextBox5.Text = Totaljawabanya



    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Totaljawabanya = TextBox5.Text
    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call tampilkandata()
        Call matikancheckbox()
        TextBox5.Enabled = False

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sqlnya = "insert into corona (LoginID,Nama,Usia,JK) values('" & Txtlogin.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "')"
        Call nyalakancheckbox()
        Call kosongkanform()
        Call jalan()
        MsgBox("Data Berhasil Tersimpan")
        Call tampilkandata()
    End Sub
End Class

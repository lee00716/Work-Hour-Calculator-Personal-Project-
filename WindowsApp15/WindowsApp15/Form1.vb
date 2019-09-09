Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'make table
        DataGridView1.Rows.Add(30)
        DataGridView2.Rows.Add(13)
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim cb2 As Integer = CInt(ComboBox2.SelectedItem)
        Dim cb3 As Integer = CInt(ComboBox3.SelectedItem)
        Dim cb4 As Integer = CInt(ComboBox4.SelectedItem)
        Dim cb5 As Integer = CInt(ComboBox5.SelectedItem)


        'Error message if the month combobox is empty
        If ComboBox1.Text = "" Then
            MessageBox.Show("Please select the month")
            Return
        End If

        'Error message if one of the comboboxes in this month's dates is empty
        If Combobox2.Text = "" Or ComboBox3.Text = "" Then
            MessageBox.Show("Please select the start date and the end date")
            Return
        End If

        'Error message if the end date is bigger than the start date (this month)
        If cb2 > cb3 Then
            MessageBox.Show("This month's start date cannot be bigger than the end date")
            Return
        End If

        'the application will skip this code unless all the values for last month are selected
        If ComboBox4.Text <> "" And ComboBox5.Text <> "" Then
            'Error message if the end date is bigger than the start date (last month)
            If cb4 > cb5 Then
                MessageBox.Show("Last month's start date cannot be bigger than the end date")
                Return
            End If
        End If



        'clear previously inserted values
        For Each row As DataGridViewRow In DataGridView1.Rows
            row.Cells.Item("ThisMonth").Value = String.Empty
        Next
        For Each row As DataGridViewRow In DataGridView2.Rows
            row.Cells.Item("DataGridViewTextBoxColumn1").Value = String.Empty
        Next
        For Each row As DataGridViewRow In DataGridView1.Rows
            row.Cells.Item("ThisDate").Value = String.Empty
        Next
        For Each row As DataGridViewRow In DataGridView2.Rows
            row.Cells.Item("DataGridViewTextBoxColumn2").Value = String.Empty
        Next


        'insert values in datagridview1 from comboboxes 
        Dim i As Integer = 0
        Dim k As Integer
        k = cb2
        Do Until i = cb3 - cb2 + 1
            DataGridView1.Rows.Item(i).Cells(1).Value = k
            DataGridView1.Rows.Item(i).Cells(0).Value = ComboBox1.Text
            i += 1
            k += 1
        Loop


        'calculates which month is the last month
        Dim Monthindex As Integer
        Dim Lastmonthindex As Integer

        Monthindex = ComboBox1.Items.IndexOf(ComboBox1.Text)

        If Monthindex = 0 Then
            Lastmonthindex = 11
        Else
            Lastmonthindex = Monthindex - 1
        End If

        'voids if nothing is selected for last month values(need this to avoid error)
        If ComboBox4.Text = "" Or ComboBox5.Text = "" Then
            Return
        Else
            If (cb5 - cb4) > 13 Then
                MessageBox.Show("Cannot add more than 14 days from last month")
                Return
            End If
            'insert values in datagridview2 from comboboxes 
            Dim o As Integer = 0
            Dim l As Integer
            l = cb4
            Do Until o = cb5 - cb4 + 1
                DataGridView2.Rows.Item(o).Cells(1).Value = l
                DataGridView2.Rows.Item(o).Cells(0).Value = ComboBox1.Items(Lastmonthindex)
                o += 1
                l += 1
            Loop
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cb2 As Integer = CInt(ComboBox2.SelectedItem)
        Dim cb3 As Integer = CInt(ComboBox3.SelectedItem)
        Dim cb4 As Integer = CInt(ComboBox4.SelectedItem)
        Dim cb5 As Integer = CInt(ComboBox5.SelectedItem)

        'Repeats error validation process in case this button is clicked first

        'Error message if the month combobox is empty
        If ComboBox1.Text = "" Then
            MessageBox.Show("Please select the month")
            Return
        End If

        'Error message if one of the comboboxes in this month's dates is empty
        If ComboBox2.Text = "" Or ComboBox3.Text = "" Then
            MessageBox.Show("Please select the start date and the end date")
            Return
        End If

        'Error message if the end date is bigger than the start date (this month)
        If cb2 > cb3 Then
            MessageBox.Show("This month's start date cannot be bigger than the end date")
            Return
        End If

        'the application will skip this code unless all the values for last month are selected
        If ComboBox4.Text <> "" And ComboBox5.Text <> "" Then
            'Error message if the end date is bigger than the start date (last month)
            If cb4 > cb5 Then
                MessageBox.Show("Last month's start date cannot be bigger than the end date")
                Return
            End If
        End If


        'clear previously calculated values
        For Each row As DataGridViewRow In DataGridView1.Rows
            row.Cells.Item("ThisMonthTotalHour").Value = String.Empty
        Next
        For Each row As DataGridViewRow In DataGridView2.Rows
            row.Cells.Item("DataGridViewTextBoxColumn11").Value = String.Empty
        Next
        For Each row As DataGridViewRow In DataGridView1.Rows
            row.Cells.Item("ThisMonthTotalMin").Value = String.Empty
        Next
        For Each row As DataGridViewRow In DataGridView2.Rows
            row.Cells.Item("DataGridViewTextBoxColumn12").Value = String.Empty
        Next



        'calculates totals in datagridview1 
        Dim i As Integer = 0
        Do Until i = cb3 - cb2 + 1
            DataGridView1.Rows.Item(i).Cells(10).Value = (((DataGridView1.Rows.Item(i).Cells(4).Value - DataGridView1.Rows.Item(i).Cells(2).Value) + (DataGridView1.Rows.Item(i).Cells(8).Value - DataGridView1.Rows.Item(i).Cells(6).Value)) * 60 + ((DataGridView1.Rows.Item(i).Cells(5).Value - DataGridView1.Rows.Item(i).Cells(3).Value) + (DataGridView1.Rows.Item(i).Cells(9).Value - DataGridView1.Rows.Item(i).Cells(7).Value))) \ 60
            DataGridView1.Rows.Item(i).Cells(11).Value = (((DataGridView1.Rows.Item(i).Cells(4).Value - DataGridView1.Rows.Item(i).Cells(2).Value) + (DataGridView1.Rows.Item(i).Cells(8).Value - DataGridView1.Rows.Item(i).Cells(6).Value)) * 60 + ((DataGridView1.Rows.Item(i).Cells(5).Value - DataGridView1.Rows.Item(i).Cells(3).Value) + (DataGridView1.Rows.Item(i).Cells(9).Value - DataGridView1.Rows.Item(i).Cells(7).Value))) Mod 60
            i += 1
        Loop


        'voids if nothing is selected for last month values(need this to avoid error)
        If ComboBox4.Text = "" Or ComboBox5.Text = "" Then
            Return
        Else
            If (cb5 - cb4) > 13 Then
                MessageBox.Show("Cannot add more than 14 days from last month")
                Return
            End If
            'calculates totals in datagridview2 
            Dim o As Integer = 0
            Do Until o = cb5 - cb4 + 1
                DataGridView2.Rows.Item(o).Cells(10).Value = (((DataGridView2.Rows.Item(o).Cells(4).Value - DataGridView2.Rows.Item(o).Cells(2).Value) + (DataGridView2.Rows.Item(o).Cells(8).Value - DataGridView2.Rows.Item(o).Cells(6).Value)) * 60 + ((DataGridView2.Rows.Item(o).Cells(5).Value - DataGridView2.Rows.Item(o).Cells(3).Value) + (DataGridView2.Rows.Item(o).Cells(9).Value - DataGridView2.Rows.Item(o).Cells(7).Value))) \ 60
                DataGridView2.Rows.Item(o).Cells(11).Value = (((DataGridView2.Rows.Item(o).Cells(4).Value - DataGridView2.Rows.Item(o).Cells(2).Value) + (DataGridView2.Rows.Item(o).Cells(8).Value - DataGridView2.Rows.Item(o).Cells(6).Value)) * 60 + ((DataGridView2.Rows.Item(o).Cells(5).Value - DataGridView2.Rows.Item(o).Cells(3).Value) + (DataGridView2.Rows.Item(o).Cells(9).Value - DataGridView2.Rows.Item(o).Cells(7).Value))) Mod 60
                o += 1
            Loop
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim cb2 As Integer = CInt(ComboBox2.SelectedItem)
        Dim cb3 As Integer = CInt(ComboBox3.SelectedItem)
        Dim cb4 As Integer = CInt(ComboBox4.SelectedItem)
        Dim cb5 As Integer = CInt(ComboBox5.SelectedItem)

        Dim thismonthtotalh As Integer = 0
        Dim thismonthtotalm As Integer = 0
        Dim lastmonthtotalh As Integer = 0
        Dim lastmonthtotalm As Integer = 0

        'accumulates each day's total (this month)
        For index As Integer = 0 To cb3 - cb2
                thismonthtotalh += Convert.ToInt32(DataGridView1.Rows(index).Cells(10).Value)
                thismonthtotalm += Convert.ToInt32(DataGridView1.Rows(index).Cells(11).Value)
            Next

        'accumulates each day's total (last month)
        For index2 As Integer = 0 To cb5 - cb4
            lastmonthtotalh += Convert.ToInt32(DataGridView2.Rows(index2).Cells(10).Value)
            lastmonthtotalm += Convert.ToInt32(DataGridView2.Rows(index2).Cells(11).Value)
        Next


        Dim finaltotalh As Integer
        Dim finaltotalm As Integer

        'adds thismonth and last month and express it in hh:mm and decimal
        finaltotalh = ((thismonthtotalh + lastmonthtotalh) * 60 + thismonthtotalm + lastmonthtotalm) \ 60
        finaltotalm = ((thismonthtotalh + lastmonthtotalh) * 60 + thismonthtotalm + lastmonthtotalm) Mod 60

        Label5.Text = finaltotalh
        Label7.Text = finaltotalm
        Label20.Text = ((thismonthtotalh + lastmonthtotalh) * 60 + thismonthtotalm + lastmonthtotalm) / 60


    End Sub
End Class

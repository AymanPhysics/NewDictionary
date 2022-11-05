Imports System.IO
Imports Microsoft.VisualBasic.Devices
Public Class Form1
    Dim dt As New DataTable With {.TableName = "Dictionary"}
    Dim dst As New DataGridViewCellStyle

    Dim Dv As New DataView
    Private Sub MySettingsBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        DG.EndEdit()
        dt.WriteXml("data.xml", XmlWriteMode.WriteSchema, True)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            dt.ReadXml("data.xml")
            Dv.Table = dt
            DG.DataSource = Dv
            DG.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
            Label1.Text = DG.Rows.Count
            DG.Columns("wav").ReadOnly = True
            DG.Columns("wav").DefaultCellStyle = (New DataGridViewLinkCell).Style
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Dim xx As String = InputBox("أدخل اسم العمود")
        If xx.Trim = "" Then Return
        dt.Columns.Add(xx)
        Dv.Table = dt
        DG.DataSource = Dv
    End Sub

    Private Sub DG_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG.CellDoubleClick
        Try
            If dt.Columns(DG.CurrentCell.ColumnIndex).DataType Is GetType(System.Byte()) Then
                Dim o As New OpenFileDialog
                o.Filter = "Images (*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF|All files (*.*)|*.*"
                If o.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim img As Image = Image.FromFile(o.FileName)
                    dt.Columns(DG.CurrentCell.ColumnIndex).ReadOnly = False
                    dt.Rows(DG.CurrentCell.RowIndex)(DG.CurrentCell.ColumnIndex) = imageToByteArray(img)
                    dt.Columns(DG.CurrentCell.ColumnIndex).ReadOnly = True
                End If
            End If

            If DG.Columns(DG.CurrentCell.ColumnIndex).Name.ToLower = "wav" Then
                Dim o As New OpenFileDialog
                o.Filter = "wav files (*.wav)|*.wav|All files (*.*)|*.*"
                If o.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim s As String = o.FileName.Split("\")(o.FileName.Split("\").Length - 1)
                    IO.File.Copy(o.FileName, Application.StartupPath & "\audio\" & s, True)
                    dt.Columns(DG.CurrentCell.ColumnIndex).ReadOnly = False
                    dt.Rows(DG.CurrentCell.RowIndex)(DG.CurrentCell.ColumnIndex) = s
                    dt.Columns(DG.CurrentCell.ColumnIndex).ReadOnly = True
                End If
            End If
        Catch
        End Try
    End Sub

    Function imageToByteArray(imageIn As System.Drawing.Image) As Byte()
        Dim ms As New MemoryStream()
        imageIn.GetThumbnailImage(100, 100, System.Drawing.Image.GetThumbnailImageAbort.Combine, System.IntPtr.Zero).Save(ms, System.Drawing.Imaging.ImageFormat.Gif)
        Return ms.ToArray()
    End Function


    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim xx As String = InputBox("أدخل اسم العمود")
        If xx.Trim = "" Then Return

        Dim d As New DataColumn With {.Caption = xx, .ColumnName = xx, .ReadOnly = True}
        d.DataType = GetType(System.Byte())
        dt.Columns.Add(d)
        dt.Columns(d.ColumnName).Caption = xx
        Dv.Table = dt
        DG.DataSource = Dv
    End Sub


    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click, ToolStripButton5.Click
        Dim xx As String = InputBox("أدخل اسم العمود")
        If xx.Trim = "" Then Return

        Dim d As New DataColumn With {.Caption = xx, .ColumnName = xx}
        d.DataType = GetType(System.Boolean)
        dt.Columns.Add(d)
        dt.Columns(d.ColumnName).Caption = xx
        Dv.Table = dt
        DG.DataSource = Dv
    End Sub

    Private Sub ToolStripButton3_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim xx As String = "Wav"
        Dim d As New DataColumn With {.Caption = xx, .ColumnName = xx}
        dt.Columns.Add(d)
        dt.Columns(d.ColumnName).Caption = xx
        Dv.Table = dt
        DG.DataSource = Dv
        DG.Columns(d.ColumnName).ReadOnly = True


    End Sub

    Private Sub DG_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG.CellClick
        Try
            If DG.Columns(DG.CurrentCell.ColumnIndex).Name.ToLower = "wav" Then
                Dim a As New Audio
                a.Play(Application.StartupPath & "\audio\" & DG.CurrentCell.Value, AudioPlayMode.Background)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        dt.Columns.RemoveAt(DG.CurrentCell.ColumnIndex)
    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        Try
            Dv.RowFilter = "([" & dt.Columns(0).ColumnName & "] like '" & IIf(RadioButton2.Checked, "%", "") & TextBox1.Text.Trim & "%' or [" & dt.Columns(0).ColumnName & "] like '%/ " & TextBox1.Text.Trim & "%' or [" & dt.Columns(0).ColumnName & "] like '%(" & TextBox1.Text.Trim & "%') and [" & dt.Columns(1).ColumnName & "] like '" & IIf(RadioButton2.Checked, "%", "") & TextBox2.Text.Trim & "%' and [" & dt.Columns(2).ColumnName & "] like '" & IIf(RadioButton2.Checked, "%", "") & TextBox3.Text.Trim & "%'"
            Label1.Text = DG.Rows.Count
        Catch ex As Exception
        End Try
    End Sub


    Private Sub DG_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DG.DataError

    End Sub

    Private Sub DG_SelectionChanged(sender As Object, e As EventArgs) Handles DG.SelectionChanged

    End Sub

    Private Sub DG_RowValidated(sender As Object, e As DataGridViewCellEventArgs) Handles DG.RowValidated
        Try

            If dt.Rows(DG.CurrentRow.Index)(0).ToString = "" Then
                dt.Rows(DG.CurrentRow.Index)(0) = "-"
            End If
            If dt.Rows(DG.CurrentRow.Index)(1).ToString = "" Then
                dt.Rows(DG.CurrentRow.Index)(1) = "-"
            End If
            If dt.Rows(DG.CurrentRow.Index)(2).ToString = "" Then
                dt.Rows(DG.CurrentRow.Index)(2) = "-"
            End If

        Catch ex As Exception

        End Try
    End Sub


End Class

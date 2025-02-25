Imports System.Data.SqlClient
Public Class Form1
    'connect to MSQL
    Dim index As Integer
    Dim conn As New SqlConnection("Server = DESKTOP-QKNE475; Database = Supermarketdb; Integrated Security = true")
    Private Function vld(ByVal ParamArray ctl() As Object) As Boolean
        'this function checks for empty field
        For i As Integer = 0 To UBound(ctl)
            If ctl(i).text = "" Then
                Return False
                Exit Function
            End If
        Next
        Return True
    End Function
#Region "ADD item detail"
    Private Sub btnadd_Click(sender As Object, e As EventArgs) Handles btnadd.Click
        'first let check validation
        If vld(txtbsn, txtbgn, txtbbn, txtbii, txtbq, txtbbno, txtbp, txtbr) = False Then
            MsgBox("Complete the form first", MsgBoxStyle.Critical)
        Else
            'check for duplication, to do this we use serial number as primary key
            conn.Open()
            Dim cmdep As SqlCommand = New SqlCommand("select id from stocktbl where id=@cosid", conn)
            cmdep.Parameters.AddWithValue("@cosid", txtbsn.Text)
            Dim rdrp As SqlDataReader = cmdep.ExecuteReader()
            If rdrp.HasRows Then
                MsgBox("Stock serial number already exist please try again!", MsgBoxStyle.Critical)
            Else
                'insert detail into db
                Dim status As Integer = 1
                Dim insrt As String = "insert into stocktbl(id,genericname,brandname,importedid,quantity,batchno,price,manufactureddate,expireddate,remark,status) values('" & txtbsn.Text & "','" & txtbgn.Text & "','" & txtbbn.Text & "','" & txtbii.Text & "','" & txtbq.Text & "','" & txtbbno.Text & "','" & txtbp.Text & "','" & DateTimePickermd.Value.ToShortDateString() & "','" & DateTimePickered.Value.ToShortDateString() & "','" & txtbr.Text & "','" & status & "')"
                InsertUpdateDelete(insrt)
                MessageBox.Show("Data successfully added", "Data Added", MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtbsn.Clear()
                txtbgn.Clear()
                txtbii.Clear()
                txtbbn.Clear()
                txtbq.Clear()
                txtbbno.Clear()
                txtbq.Clear()
                txtbp.Clear()
                txtbr.Clear()
            End If
        End If
        conn.Close()
    End Sub
#End Region
#Region "Clear Form1 Controls"
    Private Sub btnclear_Click(sender As Object, e As EventArgs) Handles btnclear.Click
        txtbsn.Clear()
        txtbgn.Clear()
        txtbii.Clear()
        txtbbn.Clear()
        txtbq.Clear()
        txtbbno.Clear()
        txtbq.Clear()
        txtbp.Clear()
        txtbr.Clear()
    End Sub
#End Region
#Region "Load Form1"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        filterdata("")
    End Sub
#End Region
#Region "View item detail to DataGridViewforitemdisplay"
    Public Sub filterdata(valuetobesearch As String)
        Try
            Dim searchq1 As String = "select * from stocktbl where CONCAT(id,genericname,brandname) like '%" & valuetobesearch & "%' and status =1"
            Dim c As New SqlCommand(searchq1, cn)
            Dim d As New SqlDataAdapter(c)
            Dim t As New DataTable()
            d.Fill(t)
            DataGridViewforitemdisplay.DataSource = t
        Catch ex As Exception
            MsgBox("Please enter the item data first", MsgBoxStyle.Information)
        End Try
    End Sub
#End Region
#Region "SEARCH Item"
    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        filterdata(txtbitemtobesearched.Text)
    End Sub
    Private Sub txtbitemtobesearched_TextChanged(sender As Object, e As EventArgs) Handles txtbitemtobesearched.TextChanged
        filterdata(txtbitemtobesearched.Text)
    End Sub
    Private Sub txtbitemtobesearched_KeyDown(sender As Object, e As KeyEventArgs) Handles txtbitemtobesearched.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnsearch_Click(sender, e)
        End If
    End Sub
#End Region
#Region "Edit item detail"
    Private Sub btnedit_Click(sender As Object, e As EventArgs) Handles btnedit.Click
        'first check the validity durind update
        If vld(txtbsn, txtbgn, txtbbn, txtbii, txtbq, txtbbno, txtbp, txtbr) = False Then
            MsgBox("Please complete the form first", MsgBoxStyle.Critical)
        Else
            'check for item data availability
            conn.Open()
            Dim c As SqlCommand = New SqlCommand("select id from stocktbl where id=@itemid", conn)
            c.Parameters.AddWithValue("@itemid", txtbsn.Text)
            Dim r As SqlDataReader = c.ExecuteReader()
            If r.HasRows Then
                Dim updateq As String = "update stocktbl set genericname='" & txtbgn.Text & "',brandname='" & txtbbn.Text & "',importedid='" & txtbii.Text & "',quantity='" & txtbq.Text & "',batchno='" & txtbbno.Text & "',price='" & txtbp.Text & "',manufactureddate='" & DateTimePickermd.Value.ToShortDateString() & "',expireddate='" & DateTimePickered.Value.ToShortDateString() & "',remark='" & txtbr.Text & "' where id='" & txtbsn.Text & "'"
                InsertUpdateDelete(updateq)
                MessageBox.Show("Successfully edited", "Data updated", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'now clear the form1 controls
                txtbsn.Clear()
                txtbgn.Clear()
                txtbii.Clear()
                txtbbn.Clear()
                txtbq.Clear()
                txtbbno.Clear()
                txtbq.Clear()
                txtbp.Clear()
                txtbr.Clear()
            Else
                MsgBox("No such serial number exist. please try again!", MsgBoxStyle.Critical)
            End If
        End If
        conn.Close()
    End Sub
    Private Sub DataGridViewforitemdisplay_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewforitemdisplay.CellContentClick
        Try
            index = e.RowIndex
            Dim selectedrow As DataGridViewRow
            selectedrow = DataGridViewforitemdisplay.Rows(index)
            '0 to 9 indicate the index of field in table
            txtbsn.Text = selectedrow.Cells(0).Value.ToString
            txtbgn.Text = selectedrow.Cells(1).Value.ToString
            txtbbn.Text = selectedrow.Cells(2).Value.ToString
            txtbii.Text = selectedrow.Cells(3).Value.ToString
            txtbq.Text = selectedrow.Cells(4).Value.ToString
            txtbbno.Text = selectedrow.Cells(5).Value.ToString
            txtbp.Text = selectedrow.Cells(6).Value.ToString
            DateTimePickermd.Text = selectedrow.Cells(7).Value.ToString
            DateTimePickered.Text = selectedrow.Cells(8).Value.ToString
            txtbr.Text = selectedrow.Cells(9).Value.ToString
        Catch ex As Exception
            MsgBox("Ctrl + C to copy item details from db or datagridview", MsgBoxStyle.Critical)

        End Try
    End Sub
#End Region
#Region "Delete item"
    Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btndelete.Click
        'check validity
        If vld(txtbsn) = False Then
            MsgBox("Please select the item from DataGridView to delete item", MsgBoxStyle.Critical)
        ElseIf IsConfirm("do you want to delete this item from database") Then
            Dim deleteq1 As String = "delete from stocktbl where id='" & txtbsn.Text & "'"
            If InsertUpdateDelete(deleteq1) Then
                MessageBox.Show("Successfully deleted", "Data Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'now clear the form1 controls
                txtbsn.Clear()
                txtbgn.Clear()
                txtbii.Clear()
                txtbbn.Clear()
                txtbq.Clear()
                txtbbno.Clear()
                txtbq.Clear()
                txtbp.Clear()
                txtbr.Clear()
            End If
        End If
    End Sub
#End Region
#Region "Refresh the form"
    Private Sub btnrefresh_Click(sender As Object, e As EventArgs) Handles btnrefresh.Click
        Me.Controls.Clear()
        InitializeComponent()
        Form1_Load(e, e)
        Refresh()
    End Sub
#End Region
End Class

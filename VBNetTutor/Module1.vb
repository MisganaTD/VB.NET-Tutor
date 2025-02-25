Imports System.Data.SqlClient

Module Module1
    'we use this all function as public in this entire solution
    Public cn As New SqlConnection("Server = DESKTOP-QKNE475; Database = Supermarketdb; Integrated Security = true")
    Public Sub Connect()
        If cn.State = ConnectionState.Closed Then cn.Open()
    End Sub

    Public Function InsertUpdateDelete(ByVal sql As String) As Boolean
        'this function for SqlCommand
        Connect()
        Dim cmd As New SqlCommand(sql, cn)
        Return cmd.ExecuteNonQuery() > 0
    End Function

    Public Function IsConfirm(ByVal message As String) As Boolean
        'function for confirm the action
        Return MessageBox.Show(message, "Confirm ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes
    End Function
    Public Function QueryAsDataTable(ByVal sql As String) As DataTable
        'function for datatable
        Dim das As New SqlDataAdapter(sql, cn)
        Dim dss As New DataSet
        das.Fill(dss, "result")
        Return dss.Tables("result")
    End Function
End Module

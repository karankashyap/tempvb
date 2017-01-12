Public Class Form1
    Private Sub addClick(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'Database1DataSet1.data' table. You can move, or remove it, as needed.
        Me.DataTableAdapter1.Fill(Me.Database1DataSet1.data)
        'TODO: This line of code loads data into the 'Database1DataSet.data' table. You can move, or remove it, as needed.
        Me.DataTableAdapter.Fill(Me.Database1DataSet.data)

        ' Set the caption bar text of the form.   
        Me.Text = "tutorialspont.com"
    End Sub
    Private Function CreateDataSet() As DataSet
        'creating a DataSet object for tables
        Dim dataset As DataSet = New DataSet()
        ' creating the student table
        Dim Students As DataTable = CreateStudentTable()
        dataset.Tables.Add(Students)
        Return dataset
    End Function
    Private Function CreateStudentTable() As DataTable
        Dim Students As DataTable
        Students = New DataTable("Student")
        ' adding columns
        AddNewColumn(Students, "System.Int32", "StudentID")
        AddNewColumn(Students, "System.String", "StudentName")
        AddNewColumn(Students, "System.String", "StudentCity")
        ' adding rows
        AddNewRow(Students, 1, "Zara Ali", "Kolkata")
        AddNewRow(Students, 2, "Shreya Sharma", "Delhi")
        AddNewRow(Students, 3, "Rini Mukherjee", "Hyderabad")
        AddNewRow(Students, 4, "Sunil Dubey", "Bikaner")
        AddNewRow(Students, 5, "Rajat Mishra", "Patna")
        Return Students
    End Function
    Private Sub AddNewColumn(ByRef table As DataTable,
   ByVal columnType As String, ByVal columnName As String)
        Dim column As DataColumn =
       table.Columns.Add(columnName, Type.GetType(columnType))
    End Sub

    'adding data into the table
    Private Sub AddNewRow(ByRef table As DataTable, ByRef id As Integer, ByRef name As String, ByRef city As String)
        Dim newrow As DataRow = table.NewRow()
        newrow("StudentID") = id
        newrow("StudentName") = name
        newrow("StudentCity") = city
        table.Rows.Add(newrow)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ds As New DataSet
        ds = CreateDataSet()
        DataGridView2.DataSource = ds.Tables("Student")



    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub
End Class

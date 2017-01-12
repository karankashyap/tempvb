Public Class Form3
    Public FI As Object
    Public Qry As String
    Public rst As New ADODB.Connection
    Public daoEngine As DAO.DBEngine

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FI = CreateObject("Busy2L16.CFixedInterface")
        Console.WriteLine(FI)

        FI.OpenDB("D:\Program Files (x86)\BusyWin", "D:\Program Files (x86)\BusyWin\DATA\", "Comp0002")

        MsgBox("Database Connected")
        Qry = "Select * from Tran1 where VchType = 9"
        'This query will return list of all the Sale vouchers as VchType=9 belong to SALE voucher in Tran1 table

        Qry = "Select * from Master1 where MasterType=6"
        'This query will return list of all the Item Masters as MasterType=6 belongs to Item Master in Master1 table.

        Debug.Print(FI)
        rst = FI.GetAllVchSeriesCodes(3)

        'rst = FI.OpenRecordset(Qry)

        Debug.Print(rst)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            rst.MoveLast
            Do While Not rst.EOF
                MsgBox(Trim$(rst!Name.Value))
                rst.MoveNext
            Loop
        End If
        FI.CloseDB

    End Sub
End Class
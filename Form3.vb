Imports System.Data.OleDb
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
Public Class Form3
    Public FI, obj As Object
    Public rst, qrst, qrst1, qrst2, qrst3 As DAO.Recordset
    Public coll As Collection
    Public bool As Boolean
    Public dbl, CurrentStock As Double
    Public dt As New DataTable
    Public objDA As New OleDbDataAdapter()
    Public ado As ADODB.Recordset

    Public Qry, str1, ErrMsg, XMLStr, VchSeries, VchDate, VchNo As String
    Public num, VchType As Integer



    Function connectDB()
        Try
            FI = CreateObject(Constant.DEFAULT_L14_DLL & "." & Constant.DEFAULT_CLASS)
        Catch
            FI = CreateObject(Constant.DEFAULT_DLL & "." & Constant.DEFAULT_CLASS)
        End Try

        FI.OpenDB(Constant.PRG_PATH, Constant.DATA_PATH, Constant.COMPANY_CODE)
        Label1.Text = "Connected to Database as: " & FI.GetCurrentUserName & " | SuperUser: " & FI.IfSuperUser(FI.GetCurrentUserName) & Constant.COMPANY_CODE
        Return FI
    End Function



    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles executeQuery.Click
        Dim ItemAlias As String = ItemAliasText.Text
        If IsDBNull(ItemAliasText) Then
            ItemAlias = False
        End If
        GetDataFromBusy("GRS", DataControl.storedQueries("GetProductInfo", ItemAlias), ItemAlias)
    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles serviceCall.Click
        MsgBox("ok")
        Dim serviceCallTest As New testService.Testing_Service

        'GetDataFromBusy("GRS", "Select * from Tran1 where VchCode = " & val)
        MsgBox(serviceCallTest.Say_Hello("oh bhai"))
    End Sub



    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub



    Public Function DebugApp(Query)
        FI = connectDB()
        Query = DataControl.storedQueries("GetProductInfo", "000001")

        qrst = FI.GetRecordset(Query)

        Dim i As Integer = 0
        Do Until i = qrst.Fields.Count
            Try
                RichTextBox1.AppendText(qrst.Fields(i).Name & " ---> " & qrst.Fields(i).Value & Environment.NewLine)
            Catch
                RichTextBox1.AppendText(qrst.Fields(i).Name & " ---> " & Environment.NewLine)
            End Try

            i += 1
        Loop

    End Function
    Public Function GetCurrentStockOfItem(ItemAlias)
        FI = connectDB()
        qrst1 = FI.GetRecordset(DataControl.storedQueries("StockStatusNew", ItemAlias))

        'Dim i As Integer = 0
        'Do Until i = qrst.Fields.Count
        'Try
        'RichTextBox1.AppendText(qrst.Fields(i).Name & " ---> " & qrst.Fields(i).Value & Environment.NewLine)
        'Catch
        'RichTextBox1.AppendText(qrst.Fields(i).Name & " ---> " & Environment.NewLine)
        'End Try
        '
        'i += 1
        'Loop

        Try
            If Not IsDBNull(qrst1()!MainOpBal.Value) And Not IsDBNull(qrst1()!MainTransBal.Value) Then
                Return qrst1()!MainOpBal.Value + qrst1()!MainTransBal.Value
            End If
        Catch
            Return False
        End Try

    End Function


    Public Function GetSTPTName(ItemAlias)
        FI = connectDB()
        qrst2 = FI.GetRecordset(DataControl.storedQueries("STPTName", ItemAlias))
        Try
            If Not IsDBNull(qrst2()!Name.Value) And Not IsDBNull(qrst2()!D1.Value) Then
                Return qrst2()!D1.Value
            End If
        Catch
            Return 1
        End Try

    End Function

    Public Function GetDataFromBusy(Method, Query, ItemAlias)

        FI = connectDB()
        If Method = "GRS" Then
            qrst = FI.GetRecordset(Query)
        ElseIf Method = "GRSBUSYDB" Then
            MsgBox("Executing... ")
            qrst = FI.GetRecordsetFromCompanyDB(DataControl.storedQueries("StockStatus", ItemAlias))
        ElseIf Method = "ExecuteQuery" Then
            qrst = FI.ExecuteQuery(DataControl.storedQueries("StockStatus", ItemAlias))
        Else
            RichTextBox1.AppendText("Query Method not defined" & Environment.NewLine)
            Return 0
        End If

        CurrentStock = GetCurrentStockOfItem(ItemAlias)
        Try

            RichTextBox2.AppendText("Current Stock:  " & CurrentStock & Environment.NewLine)

            RichTextBox2.AppendText("Item/Product: " & qrst()!PrintName.Value & Environment.NewLine)
            RichTextBox2.AppendText("MRP: " & qrst()!D2.Value & Environment.NewLine)
            RichTextBox2.AppendText("Description: " & qrst()!Address1.Value & Environment.NewLine)
            RichTextBox2.AppendText("           " & qrst()!Address2.Value & Environment.NewLine)
            RichTextBox2.AppendText("           " & qrst()!Address3.Value & Environment.NewLine)
            RichTextBox2.AppendText("           " & qrst()!Address4.Value & Environment.NewLine)
            RichTextBox2.AppendText("Code:  " & qrst()!Code.Value & Environment.NewLine)
            RichTextBox2.AppendText("Code:  " & CurrentStock & Environment.NewLine)

        Catch

        End Try

        Dim STName As String = GetSTPTName(ItemAlias)
        Dim STPTName As String = ConvertSTPT(STName)


        Dim ItemName As String
        Dim Price As Double


        ItemName = qrst()!Name.Value
        Price = qrst()!D2.Value


        Dim ItemSrNo As Integer = 1
        Dim Qty As Integer = 15
        Dim Amt As Double = Qty * Price
        VchType = 9
        VchDate = "01-04-2017"


        XMLStr = DataControl.generateXML(VchDate, STPTName, ItemName, Qty, Price, Amt)
        Try
            makeASale(VchType, XMLStr)
        Catch
            MsgBox(Constant.ERR_SALE)
        End Try


    End Function


    Public Function ConvertSTPT(STPTName)
        Dim STName As String
        If STPTName = "1%" Then
            STName = ""
        ElseIf STPTName = "12.5%" Then
            STName = "VAT/MultiTax(T)"
        ElseIf STPTName = "5%" Then
            STName = ""
        ElseIf STPTName = "Exempt" Then
            STName = ""
        ElseIf STPTName = "Services 14%" Then
            STName = ""
        Else
            STName = "VAT/MultiTax(T)"
        End If
        Return STName
    End Function
    Public Function makeASale(VchType, XMLStr)
        FI = connectDB()
        If FI.SaveVchFromXML(VchType, XMLStr, ErrMsg) Then
            RichTextBox2.AppendText("Sale Successful" & Environment.NewLine)
            Return True
        Else
            RichTextBox2.AppendText(ErrMsg & Environment.NewLine)
            RichTextBox2.AppendText(XMLStr & Environment.NewLine)
            Return False
        End If
    End Function


End Class
Imports System.Data.OleDb
Public Class Form3
    Public FI, obj As Object
    Public rst, qrst As DAO.Recordset
    Public coll As Collection
    Public bool As Boolean
    Public dbl As Double
    Public dt As New DataTable
    Public objDA As New OleDbDataAdapter()



    Public Qry, str1, ErrMsg, XMLStr, VchSeries, VchDate, VchNo As String
    Public num, VchType As Integer



    Function connectDB()
        FI = CreateObject("Busy2L16.CFixedInterface")
        FI.OpenDB("D:\Program Files (x86)\BusyWin", "D:\Program Files (x86)\BusyWin\DATA\", "Comp0002")
        Label1.Text = "Connected to Database as: " & FI.GetCurrentUserName & " | SuperUser: " & FI.IfSuperUser(FI.GetCurrentUserName)
        Return FI
    End Function



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        saveVch()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        deleteVoucher()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        PurchaseOrder()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        SaveJournalVoucher()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        DeleteJournalVoucher()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        GetDataFromBusy("GRS", "Select * from Tran1 where VchCode = 1")
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PurchaseVch()
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FI = connectDB()

        '        MsgBox("Database Connected")
        'Qry = "Select * from Tran1"
        'This query will return list of all the Sale vouchers as VchType=9 belong to SALE voucher in Tran1 table

        Qry = "Select * from Master1"
        'This query will return list of all the Item Masters as MasterType=6 belongs to Item Master in Master1 table.


        'rst = FI.GetAllVchSeriesCodes(3)
        ' ExecuteQuery only runs on DB1
        ' Select only runs with GetRecordset
        rst = FI.GetRecordset(Qry)
        num = FI.GetProvider()
        bool = FI.VchCodeExist(9)
        'str1 = FI.ExecuteQuery("INSERT INTO FAQGroup ([GrpName],[GrpCode],[GrpMasterType]) values ('TempKaranCustom', '23534', '')")
        'str1 = FI.ExecuteQuery("INSERT INTO DeletedInfo values ('Karan', 'karan','2939949')")
        'str1 = FI.ExecuteQuery("INSERT INTO StdNar values (9, 'trode','karan temp')")
        rst = FI.GetRecordset("Select * from Master1")
        bool = FI.DualTranExist()
        'dbl = FI.GetAccClosingBal(14, "16-01-2017")

        'obj = FI.GetCollectionObj()
        'addItem = FI.AddItemInCol("Sale", "2")

        'For Each Object As Object in fre 
        'RichTextBox1.Text = fre
        'Next Object

        FI.CloseDB

    End Sub


    Public Function saveVch()
        FI = connectDB()
        VchType = 9    'For SALE vch. Vchtype=9

        XMLStr = XMLStr & "<Sale>"
        XMLStr = XMLStr & "<VchSeriesName>Main</VchSeriesName><Date>01-04-2016</Date><VchType>9</VchType><VchNo>1</VchNo><STPTName>VAT/MultiTax(R)</STPTName><MasterName1>Busy Infotech Pvt. Ltd.</MasterName1><MasterName2>Main Store</MasterName2>"
        XMLStr = XMLStr & "<VchOtherInfoDetails><Narration1>Sample Narration</Narration1></VchOtherInfoDetails>"

        XMLStr = XMLStr & "<ItemEntries>"
        XMLStr = XMLStr & "<ItemDetail><SrNo>1</SrNo><ItemName>Item 1</ItemName><UnitName>Kgs.</UnitName><Qty>10</Qty><Price>105</Price><Amt>1060.5</Amt><STAmount>10.5</STAmount><STPercent>1</STPercent><TaxBeforeSurcharge>10.5</TaxBeforeSurcharge><MC>Main Store</MC></ItemDetail>"
        XMLStr = XMLStr & "<ItemDetail><SrNo>2</SrNo><ItemName>Item 2</ItemName><UnitName>Kgs.</UnitName><Qty>10</Qty><Price>100</Price><Amt>1010</Amt><STAmount>10</STAmount><STPercent>1</STPercent><TaxBeforeSurcharge>10</TaxBeforeSurcharge><MC>Main Store</MC></ItemDetail>"
        XMLStr = XMLStr & "<ItemDetail><SrNo>3</SrNo><ItemName>Item 3</ItemName><UnitName>Pcs.</UnitName><Qty>3</Qty><Price>100</Price><Amt>315</Amt><STAmount>15</STAmount><STPercent>5</STPercent><TaxBeforeSurcharge>15</TaxBeforeSurcharge><MC>Main Store</MC></ItemDetail>"
        XMLStr = XMLStr & "</ItemEntries>"

        XMLStr = XMLStr & "<BillSundries>"
        XMLStr = XMLStr & "<BSDetail><SrNo>1</SrNo><BSName>Discount</BSName><PercentVal>10</PercentVal><Amt>238.55</Amt></BSDetail>"
        XMLStr = XMLStr & "<BSDetail><SrNo>2</SrNo><BSName>Freight &amp; Forwarding Charges</BSName><Amt>100</Amt></BSDetail>"
        XMLStr = XMLStr & "</BillSundries>"
        XMLStr = XMLStr & "</Sale>"



        'Save XML string in Busy
        'Function Defination in Busy - SaveVchFromXML(ByVal p_VchType As Variant, ByVal p_XMLStr As Variant, ByRef p_ErrStr As Variant, Optional ByVal p_Modify = True) As Boolean
        If FI.SaveVchFromXML(VchType, XMLStr, ErrMsg) Then
            RichTextBox1.AppendText("Voucher Saved" & Environment.NewLine)
        Else
            RichTextBox1.AppendText(ErrMsg & Environment.NewLine)
        End If
        FI.CloseDB
    End Function


    Function deleteVoucher()
        FI = connectDB()
        VchSeries = "Main"
        VchDate = Format(CDate("10-04-2016"), "dd-mm-yyyy")  'Date shud be in 'dd-mm-yyyy' format only
        VchType = 9    'For SALE vch. Vchtype=9
        VchNo = "1"

        'Delete Voucher in Busy
        'Function Defination in Busy - DeleteVch(ByVal p_VchType As Variant, ByVal p_VchSeriesName As Variant, ByVal p_VchDate As Variant, ByVal p_VchNo As Variant, ByRef p_ErrStr As Variant) As Boolean
        If FI.DeleteVch(VchType, VchSeries, VchDate, VchNo, ErrMsg) Then
            RichTextBox1.AppendText("Voucher Deleted" & Environment.NewLine)
        Else
            RichTextBox1.AppendText(ErrMsg & Environment.NewLine)
        End If
        FI.CloseDB
    End Function


    Function PurchaseVch()
        FI = connectDB()
        VchType = 2    'For PURCHASE vch. Vchtype=2
        XMLStr = "<Purchase>"
        XMLStr = XMLStr & "<VchSeriesName>Main</VchSeriesName><Date>01-04-2016</Date><VchType>2</VchType><VchNo>1</VchNo><STPTName>VAT/Exempt</STPTName><MasterName1>Busy Infotech Pvt. Ltd.</MasterName1><MasterName2>Main Store</MasterName2>"
        XMLStr = XMLStr & "<VchOtherInfoDetails><PurchaseBillNo>Supp Purc Ref No.</PurchaseBillNo><Narration1>Sample Narration</Narration1></VchOtherInfoDetails>"

        XMLStr = XMLStr & "<ItemEntries>"
        XMLStr = XMLStr & "<ItemDetail><SrNo>1</SrNo><ItemName>Item 1</ItemName><UnitName>Kgs.</UnitName><Qty>100</Qty><Price>90</Price><Amt>9000</Amt><STAmount>90</STAmount><STPercent>1</STPercent><TaxBeforeSurcharge>90</TaxBeforeSurcharge><MC>Main Store</MC></ItemDetail>"
        XMLStr = XMLStr & "<ItemDetail><SrNo>2</SrNo><ItemName>Item 2</ItemName><UnitName>Kgs.</UnitName><Qty>100</Qty><Price>90</Price><Amt>9000</Amt><STAmount>90</STAmount><STPercent>1</STPercent><TaxBeforeSurcharge>90</TaxBeforeSurcharge><MC>Main Store</MC></ItemDetail>"
        XMLStr = XMLStr & "<ItemDetail><SrNo>3</SrNo><ItemName>Item 3</ItemName><UnitName>Pcs.</UnitName><Qty>5</Qty><Price>101</Price><Amt>530.25</Amt><STAmount>25.25</STAmount><STPercent>5</STPercent><TaxBeforeSurcharge>25.25</TaxBeforeSurcharge><MC>Main Store</MC></ItemDetail>"
        XMLStr = XMLStr & "</ItemEntries>"

        XMLStr = XMLStr & "<BillSundries>"
        XMLStr = XMLStr & "<BSDetail><SrNo>1</SrNo><BSName>Discount</BSName><PercentVal>10</PercentVal><Amt>1800</Amt></BSDetail>"
        XMLStr = XMLStr & "<BSDetail><SrNo>2</SrNo><BSName>Freight &amp; Forwarding Charges</BSName><Amt>100</Amt></BSDetail>"
        XMLStr = XMLStr & "</BillSundries>"

        XMLStr = XMLStr & "</Purchase>"



        'Save XML string in Busy
        'Function Defination in Busy - SaveVchFromXML(ByVal p_VchType As Variant, ByVal p_XMLStr As Variant, ByRef p_ErrStr As Variant, Optional ByVal p_Modify = True) As Boolean
        If FI.SaveVchFromXML(VchType, XMLStr, ErrMsg) Then
            RichTextBox1.AppendText("Purchase Voucher Saved" & Environment.NewLine)
        Else
            RichTextBox1.AppendText(ErrMsg & Environment.NewLine)
        End If
        FI.CloseDB
    End Function

    Public Function PurchaseOrder()
        FI = connectDB()
        VchType = 13    'For PURCHASE ORDER vch. Vchtype=13
        XMLStr = "<PurchaseOrder>"
        XMLStr = XMLStr & "<VchSeriesName>Main</VchSeriesName><Date>01-04-2016</Date><VchType>13</VchType><VchNo>1</VchNo><STPTName>VAT/MultiTax(T)</STPTName><MasterName1>Busy Infotech Pvt. Ltd.</MasterName1><MasterName2>Main Store</MasterName2>"

        XMLStr = XMLStr & "<ItemEntries>"
        XMLStr = XMLStr & "<ItemDetail><SrNo>1</SrNo><ItemName>Item 1</ItemName><UnitName>Pcs.</UnitName><Qty>10</Qty><QtyMainUnit>10</QtyMainUnit><QtyAltUnit>10</QtyAltUnit><Price>100</Price><Amt>1010</Amt><STAmount>10</STAmount><STPercent>1</STPercent><TaxBeforeSurcharge>10</TaxBeforeSurcharge><MC>Main Store</MC></ItemDetail>"
        XMLStr = XMLStr & "<ItemDetail><SrNo>2</SrNo><ItemName>Item 2</ItemName><UnitName>Pcs.</UnitName><Qty>20</Qty><QtyMainUnit>20</QtyMainUnit><QtyAltUnit>20</QtyAltUnit><Price>200</Price><Amt>4200</Amt><STAmount>200</STAmount><STPercent>5</STPercent><TaxBeforeSurcharge>200</TaxBeforeSurcharge><MC>Main Store</MC></ItemDetail>"
        XMLStr = XMLStr & "</ItemEntries>"

        XMLStr = XMLStr & "<BillSundries>"
        XMLStr = XMLStr & "<BSDetail><SrNo>1</SrNo><BSName>Discount</BSName><PercentVal>5</PercentVal><Amt>260.5</Amt></BSDetail>"
        XMLStr = XMLStr & "</BillSundries>"

        XMLStr = XMLStr & "<PendingOrders>"
        XMLStr = XMLStr & "<OrderDetail><MasterName1>Item 1</MasterName1><MasterName2>Busy Infotech Pvt. Ltd.</MasterName2>"
        XMLStr = XMLStr & "<OrderRefs><Method>1</Method><SrNo>1</SrNo><RefNo>1</RefNo><Date>01-04-2016</Date><DueDate>01-04-2016</DueDate><Value1>-10</Value1><Value2>-10</Value2><ItemSrNo>1</ItemSrNo><tmpMasterCode1>1228</tmpMasterCode1><tmpMasterCode2>1004</tmpMasterCode2></OrderRefs>"
        XMLStr = XMLStr & "</OrderDetail>"

        XMLStr = XMLStr & "<OrderDetail><MasterName1>Item 2</MasterName1><MasterName2>Busy Infotech Pvt. Ltd.</MasterName2>"
        XMLStr = XMLStr & "<OrderRefs><Method>1</Method><SrNo>1</SrNo><RefNo>1</RefNo><Date>01-04-2016</Date><DueDate>01-04-2016</DueDate><Value1>-20</Value1><Value2>-20</Value2><ItemSrNo>2</ItemSrNo><tmpMasterCode1>1229</tmpMasterCode1><tmpMasterCode2>1004</tmpMasterCode2></OrderRefs>"
        XMLStr = XMLStr & "</OrderDetail>"
        XMLStr = XMLStr & "</PendingOrders>"

        XMLStr = XMLStr & "</PurchaseOrder>"



        'Save XML string in Busy
        'Function Defination in Busy - SaveVchFromXML(ByVal p_VchType As Variant, ByVal p_XMLStr As Variant, ByRef p_ErrStr As Variant, Optional ByVal p_Modify = True) As Boolean
        If FI.SaveVchFromXML(VchType, XMLStr, ErrMsg) Then
            RichTextBox1.AppendText("Purchase order Saved" & Environment.NewLine)
        Else
            RichTextBox1.AppendText(ErrMsg & Environment.NewLine)
        End If
        FI.CloseDB
    End Function


    Public Function SaveJournalVoucher()
        ' Accounting Voucher
        'Build XML String of Journal Voucher to be saved in Busy
        FI = connectDB()
        XMLStr = "<Journal><VchSeriesName>Main</VchSeriesName><Date>01-04-2016</Date><VchType>16</VchType>"
        XMLStr = XMLStr & "<AccEntries>"
        XMLStr = XMLStr & "<AccDetail><SrNo>1</SrNo><AccountName>Busy Infotech Pvt. Ltd.</AccountName><AmountType>2</AmountType><AmtMainCur>5000</AmtMainCur></AccDetail>"
        XMLStr = XMLStr & "<AccDetail><SrNo>2</SrNo><AccountName>Travelling Expenses</AccountName><AmountType>1</AmountType><AmtMainCur>2000</AmtMainCur></AccDetail>"
        XMLStr = XMLStr & "<AccDetail><SrNo>3</SrNo><AccountName>Advertisement &amp; Publicity</AccountName><AmountType>1</AmountType><AmtMainCur>2000</AmtMainCur></AccDetail>"
        XMLStr = XMLStr & "<AccDetail><SrNo>4</SrNo><AccountName>Books &amp; Periodicals</AccountName><AmountType>1</AmountType><AmtMainCur>1000</AmtMainCur></AccDetail>"
        XMLStr = XMLStr & "</AccEntries>"
        XMLStr = XMLStr & "</Journal>"

        'Save XML string in Busy
        'Function Defination in Busy - SaveVchFromXML(ByVal p_VchType As Variant, ByVal p_XMLStr As Variant, ByRef p_ErrStr As Variant, Optional ByVal p_Modify = True) As Boolean
        If FI.SaveVchFromXML(16, XMLStr, ErrMsg) Then
            RichTextBox1.AppendText("Journal Voucher Saved" & Environment.NewLine)
        Else
            RichTextBox1.AppendText(ErrMsg & Environment.NewLine)
        End If

        'Close Busy Company which is opened thru OpenCSDB function
        FI.CloseDB
    End Function


    Public Function DeleteJournalVoucher()
        FI = connectDB()
        VchSeries = "Main"
        VchDate = Format(CDate("4-10-2016"), "dd-mm-yyyy")  'Date shud be in 'dd-mm-yyyy' format only
        VchType = 16    'For Journal vch. Vchtype=16
        VchNo = "1"

        'Delete Voucher in Busy
        'Function Defination in Busy - DeleteVch(ByVal p_VchType As Variant, ByVal p_VchSeriesName As Variant, ByVal p_VchDate As Variant, ByVal p_VchNo As Variant, ByRef p_ErrStr As Variant) As Boolean
        If FI.DeleteVch(VchType, VchSeries, VchDate, VchNo, ErrMsg) Then
            RichTextBox1.AppendText("Journal Voucher Deleted" & Environment.NewLine)
        Else
            RichTextBox1.AppendText(ErrMsg & Environment.NewLine)
        End If

        'Close Busy Company which is opened thru OpenCSDB function
        FI.CloseDB
    End Function

    Public Function GetDataFromBusy(Method, Query)
        FI = connectDB()
        If Method = "GRS" Then
            qrst = FI.GetRecordset(Query)
        ElseIf Method = "GRSBUSYDB" Then
            qrst = FI.GetRecordsetFromCompanyDB(Query)
        ElseIf Method = "ExecuteQuery" Then
            qrst = FI.ExecuteQuery(Query)
        Else
            RichTextBox1.AppendText("Query Method not defined" & Environment.NewLine)
            Return 0
        End If
        RichTextBox1.AppendText("Writing Data to Form" & Environment.NewLine)
        Debug.Print(qrst(0).Name)

        'Debug.Print(qrst.GetRows)
        Console.WriteLine(qrst(0))

        RichTextBox1.AppendText(qrst.Fields)

        'objDA.Fill(dt, qrst)
        'Return objDA

        'dt = qrst
        'DataGridView1.DataSource = dt
        'DataGridView1.Refresh()

    End Function
End Class
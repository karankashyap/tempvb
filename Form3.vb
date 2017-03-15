Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports System.Web
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization
Imports System.Web.Script.Serialization
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Web.Services.WebService
Public Class Form3
    Public FI, obj As Object
    Public rst, qrst, qrst1, qrst2, qrst3, qrst4 As DAO.Recordset
    Public coll As Collection
    Public bool As Boolean
    Public dbl, CurrentStock As Double
    Public dt As New DataTable
    Public objDA As New OleDbDataAdapter()
    Public ado As ADODB.Recordset
    Public random As New Random()
    Public id = random.Next(100000, 9999999)
    Public sessId = id

    Public Arr1(25, 2) As String
    Public fs As FileStream


    Public Qry, str1, ErrMsg, XMLStr, XMLStr1, XMLStr2, VchSeries, VchDate, VchNo As String
    Public num, VchType As Integer


    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim Str1XML = File.ReadAllText(opsPath & sessId & "_xml1.txt")
        Dim Str2XML = File.ReadAllText(opsPath & sessId & "_xml2.txt")
        Dim StrXML = DataControl.generateXML(VchDate, Str1XML, Str2XML)
        makeASale(9, StrXML)
    End Sub

    Public path = Constant.DATA_PATH & Constant.COMPANY_CODE & Constant.INVOICE_DIR
    Public opsPath = Constant.DATA_PATH & Constant.COMPANY_CODE & Constant.OPS_DIR



    Function connectDB()
        Try
            FI = CreateObject(Constant.DEFAULT_L14_DLL & "." & Constant.DEFAULT_CLASS)
        Catch
            FI = CreateObject(Constant.DEFAULT_DLL & "." & Constant.DEFAULT_CLASS)
        End Try
        Try
            FI.OpenDB(Constant.PRG_PATH, Constant.DATA_PATH, Constant.COMPANY_CODE)
            Label1.Text = "Connected to Database as: " & FI.GetCurrentUserName & " | SuperUser: " & FI.IfSuperUser(FI.GetCurrentUserName) & Constant.COMPANY_CODE
        Catch
            MsgBox(Constant.ERR_DBREAD)
            Close()
        End Try

        Return FI
    End Function



    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles executeQuery.Click
        Dim ItemAlias As String = ItemAliasText.Text
        If IsDBNull(ItemAliasText) Then
            ItemAlias = False
        End If
        GetDataFromBusy("GRS", DataControl.storedQueries("GetProductInfo", ItemAlias, "", "", ""), ItemAlias)
    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles serviceCall.Click
        Dim helloWorldService As New helloWorld.WebService1
        Dim val2 As String = "Microsoft"
        Dim re = helloWorldService.HelloWorld(val2, "Another String with Spaces", 734)

        'Dim serviceCallTest As New testService.Testing_Service
        'Dim re = serviceCallTest.Say_Hello("Microsoft") ', 928, 764)
        MsgBox(re)
        'GetDataFromBusy("GRS", "Select * from Tran1 where VchCode = " & val)

    End Sub


    Public Function SrvCall(url, method, params)
        Dim srv As WebRequest = WebRequest.Create(url)
        Dim resp As WebResponse = srv.GetResponse()
        Dim jsonString As String
        Using sreader As System.IO.StreamReader = New System.IO.StreamReader(resp.GetResponseStream())
            jsonString = sreader.ReadToEnd()
        End Using
        Dim jsSerializer As System.Web.Script.Serialization.JavaScriptSerializer = New System.Web.Script.Serialization.JavaScriptSerializer()
        Dim jsonData = jsSerializer.DeserializeObject(jsonString)
        'jsonData = jsSerializer.Deserialize(Of T)(jsonString)
        Using fui = System.IO.File.CreateText(opsPath & "fui_xml1.txt")
            fui.WriteLine(jsonData)
        End Using
        Console.WriteLine(jsonData)


        'srv.Method = "GET"
        'srv.ContentType = "application/x-www-form-urlencoded"
        'Dim postData As String = "test"
        'Dim byteArray As Byte() = encoding.UTF8.GetBytes(postData)
        'Dim dataStream As Stream = srv.GetRequestStream()
        'resp = srv.GetResponse()

        '        dataStream.Write(byteArray, 0, byteArray.Length)
        '       dataStream.Close()

        'MsgBox("Done")

    End Function


    Public Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim url = "http://localhost/studio/retail/retailVI/try.php"
        'SrvCall(url, "GET", "")
        FI = connectDB()


        'FILL WITH DEFAULT VALUES
        ItemQty.Text = 9
        ItemAliasText.Text = "000008"


        'CREATE INVOICE DIR AT START


        Try
            If (Not System.IO.Directory.Exists(path)) Then
                System.IO.Directory.CreateDirectory(path)
            End If
            If (Not System.IO.Directory.Exists(opsPath)) Then
                System.IO.Directory.CreateDirectory(opsPath)
            End If
        Catch
            MsgBox(Constant.ERR_PDFDIR, Title:=Constant.ERR_PDFDIR)
        End Try



        RichTextBox1.AppendText(sessId & Environment.NewLine)
        File.Create(opsPath & "\" & sessId & ".txt").Close()
        'DebugApp("")

    End Sub



    Public Function DebugApp(Query)

        Query = DataControl.storedQueries("FindBill", "112", "", "", "")

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


    Public Function PrintDebugger(Recordset)
        Dim i As Integer = 0
        Do Until i = Recordset.Fields.Count
            Try
                RichTextBox1.AppendText(Recordset.Fields(i).Name & " ---> " & Recordset.Fields(i).Value & Environment.NewLine)
            Catch
                RichTextBox1.AppendText(Recordset.Fields(i).Name & " ---> " & Environment.NewLine)
            End Try
            i += 1
        Loop
    End Function


    Public Function GetMaterialCentres(itemMatCenter)
        If itemMatCenter Then
            qrst4 = FI.GetRecordset(DataControl.storedQueries("MatCentre", itemMatCenter, "", "", ""))
        Else
            qrst4 = FI.GetRecordset(DataControl.storedQueries("MatCentre", "", "", "", ""))
        End If

    End Function



    Public Function GetCurrentStockOfItem(ItemAlias)


        'qrst1 = FI.GetRecordset(DataControl.storedQueries("StockStatusNew", ItemAlias))
        qrst1 = FI.GetRecordset(DataControl.storedQueries("StockStatusNew", ItemAlias, "", "", ""))
        If Constant.CURRENT_MODE = "DEBUG" Then
            Dim i = 0
            Try
                Do Until i = qrst1.Fields.Count

                    Try
                        RichTextBox1.AppendText(qrst1.Fields(i).Name & " ---> " & qrst1.Fields(i).Value & Environment.NewLine)
                    Catch
                        RichTextBox1.AppendText(qrst1.Fields(i).Name & " ---> " & Environment.NewLine)
                    End Try
                    i += 1
                Loop

            Catch
            End Try
        End If


        Try
            If Not IsDBNull(qrst1()!MainOpBal.Value) And Not IsDBNull(qrst1()!MainTransBal.Value) Then
                Return qrst1()!MainOpBal.Value + qrst1()!MainTransBal.Value
            End If
        Catch
            Return False
        End Try

    End Function


    Public Function GetSTPTName(ItemAlias)

        qrst2 = FI.GetRecordset(DataControl.storedQueries("STPTName", ItemAlias, "", "", ""))
        Try
            If Not IsDBNull(qrst2()!Name.Value) And Not IsDBNull(qrst2()!D1.Value) Then
                Return qrst2()!D1.Value
            End If
        Catch
            Return 1
        End Try

    End Function

    Public Function GetDataFromBusy(Method, Query, ItemAlias)


        If Method = "GRS" Then
            qrst = FI.GetRecordset(Query)
        ElseIf Method = "GRSBUSYDB" Then
            MsgBox("Executing... ")
            qrst = FI.GetRecordsetFromCompanyDB(DataControl.storedQueries("StockStatus", ItemAlias, "", "", ""))
        ElseIf Method = "ExecuteQuery" Then
            qrst = FI.ExecuteQuery(DataControl.storedQueries("StockStatus", ItemAlias, "", "", ""))
        Else
            RichTextBox1.AppendText("Query Method not defined" & Environment.NewLine)
            Return 0
        End If

        CurrentStock = GetCurrentStockOfItem(ItemAlias)

        Dim STName As String = GetSTPTName(ItemAlias)
        Dim STPTName As String = ConvertSTPT(STName)
        Dim ItemName As String
        Dim Price As Double

        ItemName = qrst()!Name.Value
        Price = qrst()!D2.Value

        Dim ItemSrNo As Integer = 1
        'TO-DO Quantity 
        Dim Qty = ItemQty.Text
        Dim Amt As Double = Qty * Price
        VchType = Constant.VCH_TYPE
        VchDate = Constant.FY_DATE

        If Constant.CURRENT_MODE = "DEBUG" Then
            PrintDebugger(qrst)
            Try
                RichTextBox2.AppendText("Current Stock:  " & CurrentStock & Environment.NewLine)
                RichTextBox2.AppendText("Item/Product: " & qrst()!PrintName.Value & Environment.NewLine)
                RichTextBox2.AppendText("MRP: " & qrst()!D2.Value & Environment.NewLine)
                RichTextBox2.AppendText("Description: " & qrst()!Address1.Value & Environment.NewLine)
                RichTextBox2.AppendText("           " & qrst()!Address2.Value & Environment.NewLine)
                RichTextBox2.AppendText("           " & qrst()!Address3.Value & Environment.NewLine)
                RichTextBox2.AppendText("           " & qrst()!Address4.Value & Environment.NewLine)
                RichTextBox2.AppendText("Code:  " & qrst()!Code.Value & Environment.NewLine)
                RichTextBox2.AppendText("Closing Stock:  " & CurrentStock & Environment.NewLine)
                RichTextBox2.AppendText("Amount: " & Amt & Environment.NewLine)
                RichTextBox2.AppendText("STPT: " & STPTName & Environment.NewLine)
            Catch
            End Try
        End If

        XMLStr = DataControl.XMLItemDetail(ItemName, Qty, Price, Amt, STPTName, sessId)



        'generatePDF(49)

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



    Public Function grap()
        Try
            makeASale(VchType, XMLStr)
        Catch
            MsgBox(Constant.ERR_SALE)
        End Try
    End Function


    Public Function makeASale(VchType, XMLStr)

        Try
            Dim LastBillDetails As DAO.Recordset = FI.GetRecordset(DataControl.storedQueries("LastBill", "", "", "", ""))
            'Add Quantity as well
            Dim LastBillNo = LastBillDetails()!VchNo.Value
            Dim LastBillBase = LastBillDetails()!VchAmtBaseCur.Value
            Dim LastBillFinal = LastBillDetails()!VchSalePurcAmt.Value

            RichTextBox2.AppendText("Last Bill No.: " & LastBillNo & Environment.NewLine)
            If FI.SaveVchFromXML(VchType, XMLStr, ErrMsg) Then
                RichTextBox2.AppendText("Sale Successful" & Environment.NewLine)
                RichTextBox2.AppendText("This Bill No.: " & LastBillNo + 1 & Environment.NewLine)

                Dim ValidateCurrBill = ValidateBillNo(LastBillNo, LastBillBase, LastBillFinal)

                If (ValidateCurrBill = True) Then
                    generatePDF(LastBillNo + 1)
                Else
                    Dim CurrBillDetails As DAO.Recordset = FI.GetRecordset(DataControl.storedQueries("LastBill", "", "", "", ""))
                    Dim CurrBillNo = CurrBillDetails()!VchNo.Value
                    Dim CurrBillBase = CurrBillDetails()!VchAmtBaseCur.Value
                    Dim CurrBillFinal = CurrBillDetails()!VchSalePurcAmt.Value
                    generatePDF(CurrBillNo + 1)

                End If



                Return True
            End If
        Catch
            MsgBox(Constant.ERR_SALE)
            If Constant.CURRENT_MODE = "DEBUG" Then
                RichTextBox2.AppendText(ErrMsg & Environment.NewLine)
                RichTextBox2.AppendText(XMLStr & Environment.NewLine)
            End If
            Close()
        End Try
        Return False

    End Function


    Public Function ValidateBillNo(BillNo, BillBase, BillFinal)
        Dim BillDetails1 As DAO.Recordset = FI.GetRecordset(DataControl.storedQueries("LastBill", "", "", "", ""))
        Dim BillNo1 = BillDetails1()!VchNo.Value
        Dim BillBase1 = BillDetails1()!VchAmtBaseCur.Value
        Dim BillFinal1 = BillDetails1()!VchSalePurcAmt.Value

        If BillNo = BillNo1 - 1 Then
            'And BillBase = BillBase1 And BillFinal = BillFinal1 Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Function generatePDF(VchCode)
        Dim pdfGen = False
        Dim InvoicePath = Constant.DATA_PATH & Constant.COMPANY_CODE & Constant.INVOICE_DIR & VchCode

        Try
            pdfGen = FI.GeneratePDFForInvoice(VchCode + 5, InvoicePath)
        Catch
            MsgBox(Constant.ERR_PDF)
        End Try
        Try
            System.Diagnostics.Process.Start(InvoicePath & ".pdf")
            Dim Reddd = InvoicePath & ".pdf"
        Catch
            MsgBox(Constant.ERR_OPNPDF)
        End Try

        Return pdfGen

    End Function






End Class
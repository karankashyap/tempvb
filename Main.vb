Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports System.Windows.Forms

Public Class MainApp
    Public FI, obj As Object
    Public rst, qrst, qrst1, qrst2, qrst3, qrst4 As DAO.Recordset
    Public coll As Collection
    Public bool As Boolean
    Public dbl, CurrentStock As Double
    Public dt, dk As New DataTable
    Public ds As DataSet = New DataSet
    Public objDA As New OleDbDataAdapter()
    Public ado As ADODB.Recordset
    Public random As New Random()
    Public id = random.Next(100000, 9999999)
    Public sessIdFromVB = id

    Public Arr1(25, 2) As String
    Public fs As FileStream

    Public Qry, str1, ErrMsg, XMLStr, XMLStr1, XMLStr2, VchSeries, VchDate, VchNo As String
    Public num, VchType, Qty As Integer


    Public CurrentDir = File.ReadAllText("C:\RetailVI_setVal.txt")

    Public PRG_PATH = File.ReadAllText(CurrentDir + "/settings/" & "program_path.txt")
    Public DATA_PATH = File.ReadAllText(CurrentDir + "/settings/" & "data_path.txt")
    Public WEB_SERVICE_PATH = File.ReadAllText(CurrentDir + "/settings/" & "web_service_path.txt")
    Public COMPANY_CODE = File.ReadAllText(CurrentDir + "/settings/" & "company_code.txt")


    Public path = DATA_PATH & COMPANY_CODE & Constant.INVOICE_DIR
    Public opsPath = DATA_PATH & COMPANY_CODE & Constant.OPS_DIR

    Public Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub




    Public Function Checkout(sessId, userId)
        If System.IO.File.Exists(opsPath & sessId & "_" & userId & "_bill.txt") Then
            Return File.ReadAllText(opsPath & sessId & "_" & userId & "_bill.txt")
        End If

        Dim sessFile = opsPath & "\" & sessId & ".txt"
        Dim reader As System.IO.StreamReader =
        System.IO.File.OpenText(sessFile)
        Console.WriteLine(sessFile)
        Console.WriteLine(reader)
        Dim ItemAlias
        Dim o = 0
        Try
            Do
                ItemAlias = reader.ReadLine()
                If ItemAlias Is Nothing Then
                    Exit Do
                End If
                If ItemAlias = "xxxxxx" Then
                    Continue Do
                End If

                Dim Price = File.ReadAllText(opsPath & sessId & "_" & ItemAlias & "p.txt")
                Dim ItemName = File.ReadAllText(opsPath & sessId & "_" & ItemAlias & "n.txt")
                Dim Qty = File.ReadAllText(opsPath & sessId & "_" & ItemAlias & "q.txt")
                Dim ItemDesc = File.ReadAllText(opsPath & sessId & "_" & ItemAlias & "d.txt")
                Convert.ToInt32(Qty)
                If Qty <= 0 Then
                    Continue Do
                End If
                Dim STPTName = File.ReadAllText(opsPath & sessId & "_" & ItemAlias & "s.txt")
                Dim Amt = Price * Qty
                XMLStr = DataControl.XMLItemDetail(ItemName, Qty, Price, Amt, STPTName, sessId)

            Loop Until ItemAlias Is Nothing
        Catch

        End Try
        Try
            Dim Str1XML = File.ReadAllText(opsPath & sessId & "_xml1.txt")
            Dim Str2XML = File.ReadAllText(opsPath & sessId & "_xml2.txt")
            Dim StrXML = DataControl.generateXML(Constant.FY_DATE, Str1XML, Str2XML)
            Return makeASale(9, StrXML, userId, sessId)
        Catch
            Return "success=false&sale=Error&msg=Error in billing. Sale did not proceed."
        End Try
    End Function


    Function connectDB()
        Try
            FI = CreateObject(Constant.DEFAULT_L14_DLL & "." & Constant.DEFAULT_CLASS)
        Catch
            FI = CreateObject(Constant.DEFAULT_DLL & "." & Constant.DEFAULT_CLASS)
        End Try
        Try
            FI.OpenDB(PRG_PATH, DATA_PATH, COMPANY_CODE)
            '            Label1.Text = "Connected to Database as: " & FI.GetCurrentUserName & " | SuperUser: " & FI.IfSuperUser(FI.GetCurrentUserName) & Constant.COMPANY_CODE
        Catch
            MsgBox(Constant.ERR_DBREAD)
            Close()
        End Try

        Return FI
    End Function



    Public Function AddItemToCart(ItemAlias, sessId, userId)
        Dim BillFile = opsPath & "\" & sessId & "_" & userId & "_bill.txt"
        If (System.IO.File.Exists(BillFile)) Then
            Dim Sale = File.ReadAllText(BillFile)
            Return "success=false&msg=Bill Already Generated for this session id&" & Sale
        End If
        If (System.IO.File.Exists(opsPath & "\" & sessId & ".txt")) Then

        Else
            CreateNewSessionFile(sessId)
        End If
        Return GetDataFromBusy("GRS", DataControl.storedQueries("GetProductInfo", ItemAlias, "", "", ""), ItemAlias, sessId, userId)
    End Function

    Public Function CreateNewSessionFile(sessId)
        Try
            File.Create(opsPath & "\" & sessId & ".txt").Close()
        Catch
        End Try
        Return True

    End Function



    Public Function editItemQuantity(ItemAlias, mode, sessId)
        Dim newValue = 0
        Try
            newValue = File.ReadAllText(opsPath & sessId & "_" & ItemAlias & "q.txt")
        Catch
        End Try

        If (mode = "add") Then
            newValue = newValue + 1
        Else
            newValue = newValue - 1
        End If

        If (newValue = -1) Then
            DeleteItem(ItemAlias, sessId)
            Return "quantity=-1&product=removed"
        ElseIf (newValue = 0) Then
            DeleteItem(ItemAlias, sessId)
            Return "quantity=-1&product=removed"
        Else
        End If

        SaveValue(ItemAlias, "q", newValue, sessId)
        Return "quantity=" & newValue
    End Function

    Public Function SrvCall(url, method, params)
        Dim srv As WebRequest = WebRequest.Create(WEB_SERVICE_PATH & url & ".php?" & params)
        Dim resp As WebResponse = srv.GetResponse()
        Dim dataStream As Stream = resp.GetResponseStream()
        Dim Reader As StreamReader = New StreamReader(dataStream)
        Dim strData As String = Reader.ReadToEnd()

        MsgBox(strData)

        Return strData



    End Function


    Public Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FI = connectDB()
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

        Try
            HttpListener.Main()
        Catch
            HttpListener.Main()
        End Try


    End Sub




    Public Function GetMaterialCentres(itemMatCenter)
        qrst4 = FI.GetRecordset(DataControl.storedQueries("MatCentre", "", "", "", ""))
        Dim i As Integer = 0
        Dim params = ""
        Do Until i = qrst4.Fields.Count
            Try
                params = qrst4()!MCName.Value & "=" & qrst4()!MasterCode2.Value & "&"
                params = params & qrst4()!PGName.Value & "=" & qrst4()!PGCode.Value & "&"
            Catch
                params = "Main Store=201"
            End Try
            i += 1
        Loop

        Return params

        'End If


    End Function



    Public Function GetCurrentStockOfItem(ItemAlias)
        Dim CurrentStore = 201 'temp TO DO TODO
        qrst1 = FI.GetRecordset(DataControl.storedQueries("StockStatusNew", ItemAlias, CurrentStore, "", ""))

        Try
            If Not IsDBNull(qrst1()!MainOpBal.Value) And Not IsDBNull(qrst1()!MainTransBal.Value) Then
                Return qrst1()!MainOpBal.Value + qrst1()!MainTransBal.Value
            End If
        Catch
            Return False
        End Try
        Return False
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
        Return 1
    End Function

    Public Function lol(msg)
        Return msg
    End Function

    Public Function GetDataFromBusy(Method, Query, ItemAlias, sessId, userId)

        If Method = "GRS" Then
            qrst = FI.GetRecordset(Query)
        ElseIf Method = "GRSBUSYDB" Then
            MsgBox("Executing... ")
            qrst = FI.GetRecordsetFromCompanyDB(DataControl.storedQueries("StockStatus", ItemAlias, "", "", ""))
        ElseIf Method = "ExecuteQuery" Then
            qrst = FI.ExecuteQuery(DataControl.storedQueries("StockStatus", ItemAlias, "", "", ""))
        Else

            Return 0
        End If

        CurrentStock = GetCurrentStockOfItem(ItemAlias)

        Dim STName As String = GetSTPTName(ItemAlias)
        Dim STPTName As String = ConvertSTPT(STName)
        Dim ItemName As String
        Dim Price As Double

        ItemName = qrst()!Name.Value
        Price = qrst()!D2.Value
        Dim ItemDesc As String = ""
        Try
            ItemDesc = qrst()!Address1.Value & " "
            ItemDesc = ItemDesc & qrst()!Address2.Value & " "
            ItemDesc = ItemDesc & qrst()!Address3.Value & " "
            ItemDesc = ItemDesc & qrst()!Address4.Value
        Catch
        End Try
        Dim ItemSrNo As Integer = 1
        'TO-DO Quantity from Android

        If Qty = 0 Then
            Qty = 1
        Else
            Try
                Qty = File.ReadAllText(opsPath & sessId & "_" & ItemAlias & "q.txt")
            Catch
                Qty = 1
            End Try
        End If
        SaveValue(ItemAlias, "q", Qty, sessId)
        SaveValue(ItemAlias, "p", Price, sessId)
        SaveValue(ItemAlias, "n", ItemName, sessId)
        SaveValue(ItemAlias, "s", STPTName, sessId)
        SaveValue(ItemAlias, "d", ItemDesc, sessId)
        SaveValue(ItemAlias, "u", userId, sessId)

        Dim Amt As Double = Qty * Price
        VchType = Constant.VCH_TYPE
        VchDate = Constant.FY_DATE

        Using addItemAlias = System.IO.File.AppendText(opsPath & "\" & sessId & ".txt")
            addItemAlias.WriteLine(ItemAlias)
        End Using

        Dim param = "fn=addToCart&" & "alias=" & ItemAlias & "&ItemName=" & ItemName & "&quantity=" & Qty & "&Price=" & Price & "&Amt=" & Amt & "&CurrentStock=" & CurrentStock & "&STPTName=" & STPTName & "&sessId=" & sessId & "&desc=" & ItemDesc & "&userId=" & userId

        Return param





    End Function



    Public Function SaveValue(ItemAlias, what, howMuch, sessId)
        Using ItemQty = System.IO.File.CreateText(opsPath & sessId & "_" & ItemAlias & what & ".txt")
            ItemQty.WriteLine(howMuch)
        End Using
        Return True
    End Function


    Public Function DeleteItem(ItemAlias, sessId)
        Dim pathDel = opsPath & sessId & "_" & ItemAlias
        Try
            File.Delete(pathDel & "d.txt")
            File.Delete(pathDel & "n.txt")
            File.Delete(pathDel & "p.txt")
            File.Delete(pathDel & "q.txt")
            File.Delete(pathDel & "s.txt")
            RemoveLines(sessId, ItemAlias)
            Return "success=true&msg=Item Removed"
        Catch
            Return "success=false&msg=Item could not be Removed."
        End Try
    End Function

    Private Sub RemoveLines(sessId, ItemAlias)
        Dim FilePath As String = opsPath & sessId & ".txt"
        File.WriteAllText(FilePath, File.ReadAllText(FilePath).Replace(ItemAlias, "xxxxxx"))

    End Sub


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




    Public Function makeASale(VchType, XMLStr, userId, sessId)
        Dim ret As String
        Dim LastBillNo, LastBillBase, LastBillFinal
        Try
            Console.WriteLine("First Try")
            Dim LastBillDetails As DAO.Recordset = FI.GetRecordset(DataControl.storedQueries("LastBill", "", "", "", ""))
            LastBillNo = LastBillDetails()!VchNo.Value
            LastBillBase = LastBillDetails()!VchAmtBaseCur.Value
            LastBillFinal = LastBillDetails()!VchSalePurcAmt.Value
        Catch
            Console.WriteLine("First Catch")
        End Try

        Try
            Console.WriteLine("Secnond Try")
            Dim CurrBillNo
            If FI.SaveVchFromXML(VchType, XMLStr, ErrMsg) Then
                Dim ValidateCurrBill = ValidateBillNo(LastBillNo, LastBillBase, LastBillFinal)
                If (ValidateCurrBill = True) Then
                    Dim pdf = False 'generatePDF(LastBillNo + 1)
                    Console.WriteLine("Val1")
                    ret = "success=true&sale=" & LastBillNo + 1 & "&pdf=" & pdf & "&userId=" & userId
                    SaveBillInfo(sessId, LastBillNo, userId)
                Else
                    Dim CurrBillDetails As DAO.Recordset = FI.GetRecordset(DataControl.storedQueries("LastBill", "", "", "", ""))
                    CurrBillNo = CurrBillDetails()!VchNo.Value
                    Dim CurrBillBase = CurrBillDetails()!VchAmtBaseCur.Value
                    Dim CurrBillFinal = CurrBillDetails()!VchSalePurcAmt.Value
                    Dim pdf = False 'generatePDF(CurrBillNo + 1)
                    Console.WriteLine("Val2")
                    ret = "success=true&sale=" & CurrBillNo + 1 & "&pdf=" & pdf & "&userId=" & userId
                    SaveBillInfo(sessId, LastBillNo, userId)
                End If
            Else
                ret = "success=false&sale=Error&msg=" & ErrMsg
            End If
        Catch
            ret = "success=false&sale=Error&msg=Error in Billing. Invalid Bill Number. Code: 2MAS"

        End Try

        Console.WriteLine("End of Try Blocks")
        Return ret

    End Function


    Public Function SaveBillInfo(sessId, BillNo, UserId)
        BillNo = BillNo.Replace(" ", "")
        BillNo = BillNo + 1
        Using ItemQty = System.IO.File.CreateText(opsPath & sessId & "_" & UserId & "_bill.txt")
            ItemQty.WriteLine("sessId=" & sessId & "&sale=" & BillNo & "&userId=" & UserId)
        End Using

        Using BILLS = System.IO.File.AppendText(opsPath & "\BILLS.txt")
            BILLS.WriteLine(sessId & vbTab & BillNo & vbTab & UserId & vbTab & DateTime.Now.ToString() & vbNewLine)
        End Using
        'Dim XLA As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        'Dim XLW As Excel.Workbook
        'Dim XLS As Excel.Worksheet
        Return True
    End Function


    Public Function CheckBilling(sessId, userId)
        Try
            Return File.ReadAllText(opsPath & sessId & "_" & userId & "_bill.txt")
        Catch
            Return "sessId=" & sessId & "&sale=Error" & "&msg=Sale was not recorded. Checkout again."
        End Try


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

        Dim InvoicePath = DATA_PATH & COMPANY_CODE & Constant.INVOICE_DIR & VchCode

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



    Public Function FileCleanup(token)
        Dim directory = opsPath
        Dim oldDir = opsPath & "old\"
        For Each filename As String In IO.Directory.GetFiles(directory, "**", IO.SearchOption.AllDirectories)
            My.Computer.FileSystem.MoveFile(filename, oldDir & filename)
        Next
        Return True
    End Function









End Class
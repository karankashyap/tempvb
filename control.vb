
Public Class DataControl
    Inherits Form3
    'Public funcName, VchNo, VchType, VchName, VchDate, ItemSrNo, ItemName, UnitName, Qty, Price, Amt, STAmt, STPerc As String

    Public Shared Function storedQueries(qName)
        If qName = "StockStatus" Then
            Return "SELECT M.Name AS Name,(Select Top 1 NameAlias from Help1 as H1 where 
                    H1.NameOrAlias = 1 and H1.Code = M.PG) AS ParentGrpName, M.Mc AS McName, 
                    S1.MTB AS MainTransBal, S1.ATB AS AltTransBal,(Select Top 1 NameAlias from Help1 as 
                    H1 where H1.NameOrAlias = 1 and H1.Code = M.U1) as MU, (Select Top 1 NameAlias 
                    from Help1 as H1 where H1.NameOrAlias = 1 and H1.Code = M.U2) as AU, (select 
                    Sum(d1) from tran4 where tran4.mastercode1=  M.C and  tran4.mastercode2= M.Cm) 
                    AS MainOpBal, (select Sum(d2) from tran4 where tran4.mastercode1=  M.C and  
                    tran4.mastercode2= M.Cm) AS AltOPBal,(select Sum(d3) from tran4 where 
                    tran4.mastercode1=  M.C and  tran4.mastercode2= M.Cm) AS AmtOpBal FROM (Select 
                    A.Name as Name,A.Alias as Alias,A.PrintName as PrintName,A.code as C,A.ParentGrp 
                    as PG,A.I4 as I4,A.D9 as D9,A.CM1 as U1,A.CM2 as U2,B.Name as Mc,B.Code as CM 
                    ,A.D2  from Master1 as A,Master1 as B where A.Mastertype=6 and B.Mastertype=11) AS 
                    M LEFT JOIN (SELECT mastercode1, Mastercode2, sum(value1) AS MTB,sum(value2) AS 
                    ATB From tran2  Where rectype = 2  group by Mastercode1,Mastercode2) AS S1 ON 
                    (S1.Mastercode1 = M.c) AND (S1.Mastercode2 = M.CM)   ORDER BY M.Name "
        End If

    End Function


    Public Shared Function generateXML()

        Dim funcName = "1"
        Dim VchNo = "1"
        Dim VchName = " 1"
        Dim VchDate = "1"
        Dim ItemSrNo = "1"
        Dim ItemName = "1"
        Dim UnitName = "1"
        Dim Qty = "1"
        Dim Price = "1"
        Dim Amt = "1"
        Dim STAmt = "1"
        Dim STPerc = "1"
        'Dim VchType = VchType    'For SALE vch. Vchtype=9
        Dim VchType = 9

        'Sale,01-04-2016,
        Dim XmlStr As String
        XmlStr = XmlStr & "<" & VchName & ">"
        XmlStr = XmlStr & "<VchSeriesName>Main</VchSeriesName><Date>" & VchDate & "</Date><VchType>" & VchType & "</VchType><VchNo>" & VchNo & "</VchNo><STPTName>VAT/MultiTax(R)</STPTName><MasterName1>Busy Infotech Pvt. Ltd.</MasterName1><MasterName2>Main Store</MasterName2>"
        XmlStr = XmlStr & "<VchOtherInfoDetails><Narration1>Sample Narration</Narration1></VchOtherInfoDetails>"

        XmlStr = XmlStr & "<ItemEntries>"
        XmlStr = XmlStr & "<ItemDetail><SrNo>" & ItemSrNo & "</SrNo><ItemName>" & ItemName & "</ItemName><UnitName>" & UnitName & "</UnitName><Qty>" & Qty & "</Qty><Price>" & Price & "</Price><Amt>" & Amt & "</Amt><STAmount>" & STAmt & "</STAmount><STPercent>" & STPerc & "</STPercent><TaxBeforeSurcharge>10.5</TaxBeforeSurcharge><MC>Main Store</MC></ItemDetail>"
        XmlStr = XmlStr & "</ItemEntries>"

        XmlStr = XmlStr & "<BillSundries>"
        XmlStr = XmlStr & "<BSDetail><SrNo>1</SrNo><BSName>Discount</BSName><PercentVal>10</PercentVal><Amt>238.55</Amt></BSDetail>"
        XmlStr = XmlStr & "<BSDetail><SrNo>2</SrNo><BSName>Freight &amp; Forwarding Charges</BSName><Amt>100</Amt></BSDetail>"
        XmlStr = XmlStr & "</BillSundries>"
        XmlStr = XmlStr & "</Sale>"

        MsgBox(XmlStr)

    End Function
End Class
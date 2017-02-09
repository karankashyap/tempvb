
Public Class DataControl
    Inherits Form3


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
        'funcName, VchNo, VchType, VchName, VchDate, ItemSrNo, ItemName, UnitName, Qty, Price, Amt, STAmt, STPerc, TaxBfrSurChg,MC

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
        Dim TaxBfrSurChg, MC, BSSrNo1, BSName1, BSPercentVal, BSAmt, BSSrNo2, BSName2, BSAmt2 As String
        'BSSrNo1, BSName1, BSPercentVal = "1"
        'BSAmt, BSSrNo2, BSName2, BSAmt2 = "1"

        'Sale,01-04-2016,
        Dim XmlStr As String
        XmlStr = XmlStr & "<" & VchName & ">"
        XmlStr = XmlStr & "<VchSeriesName>Main</VchSeriesName><Date>" & VchDate & "</Date><VchType>" & VchType & "</VchType><VchNo>" & VchNo & "</VchNo><STPTName>VAT/MultiTax(R)</STPTName><MasterName1>Busy Infotech Pvt. Ltd.</MasterName1><MasterName2>Main Store</MasterName2>"
        XmlStr = XmlStr & "<VchOtherInfoDetails><Narration1>Sample Narration</Narration1></VchOtherInfoDetails>"

        XmlStr = XmlStr & "<ItemEntries>"
        XmlStr = XmlStr & "<ItemDetail><SrNo>" & ItemSrNo & "</SrNo><ItemName>" & ItemName & "</ItemName><UnitName>" & UnitName & "</UnitName><Qty>" & Qty & "</Qty><Price>" & Price & "</Price><Amt>" & Amt & "</Amt><STAmount>" & STAmt & "</STAmount><STPercent>" & STPerc & "</STPercent><TaxBeforeSurcharge>" & TaxBfrSurChg & "</TaxBeforeSurcharge><MC>" & MC & "</MC></ItemDetail>"
        XmlStr = XmlStr & "</ItemEntries>"

        XmlStr = XmlStr & "<BillSundries>"
        XmlStr = XmlStr & "<BSDetail><SrNo>" & BSSrNo1 & "</SrNo><BSName>" & BSName1 & "</BSName><PercentVal>" & BSPercentVal & "</PercentVal><Amt>" & BSAmt & "</Amt></BSDetail>"
        XmlStr = XmlStr & "<BSDetail><SrNo>" & BSSrNo2 & "</SrNo><BSName>" & BSName2 & "</BSName><Amt>" & BSAmt2 & "</Amt></BSDetail>"
        XmlStr = XmlStr & "</BillSundries>"
        XmlStr = XmlStr & "</Sale>"

        Return XmlStr

    End Function


    Public Shared Function generatePurchaseXML()
        Dim XMLStr = "<Purchase>"
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
    End Function
End Class
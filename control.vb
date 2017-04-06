
Public Class DataControl
    Inherits MainApp

    Public Shared Function storedQueries(qName, ItemAlias, param3, param4, param5)
        Dim RetQrr As String = ""
        If qName = "StockStatus" Then
            RetQrr = "SELECT M.Name AS Name,(Select Top 1 NameAlias from Help1 as H1 where 
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
        ElseIf qName = "StockStatusNew" Then
            RetQrr = "Select M.Name As Name,M.Alias As Alias,M.Printname As PrintName ,M.D9 As SVP,M.I4 As SVM,M.C As Code,M.PG As ParentGrp,M.D2,
                        (Select Top 1 NameAlias from Help1 as H1 where H1.NameOrAlias = 1 And H1.Code = M.PG) AS ParentGrpName,
                        M.CM as MCCode, M.Mc As McName, S1.MTB As MainTransBal, S1.ATB As AltTransBal,
                        (Select Top 1 NameAlias from Help1 as H1 where H1.NameOrAlias = 1 And H1.Code = M.U1) as MU,
                        (Select Top 1 NameAlias from Help1 as H1 where H1.NameOrAlias = 1 And H1.Code = M.U2) as AU,
                        (select Sum(d1) from tran4 where tran4.mastercode1=  M.C And  tran4.mastercode2= M.Cm) AS MainOpBal,
                        (select Sum(d2) from tran4 where tran4.mastercode1=  M.C And  tran4.mastercode2= M.Cm) AS AltOPBal,
                        (select Sum(d3) from tran4 where tran4.mastercode1=  M.C And  tran4.mastercode2= M.Cm) AS AmtOpBal
                        FROM(Select A.Name as Name, A.Alias As Alias, A.PrintName As PrintName, A.code As C, A.ParentGrp As PG, A.I4 As I4,
                        A.D9 as D9, A.CM1 As U1, A.CM2 As U2, B.Name As Mc, B.Code As CM , A.D2  from Master1 As A, Master1 as B
                        where A.Mastertype = 6 And B.Mastertype = 11 And B.Code = " & param3 & "And 
                        A.Alias ='" & ItemAlias & "') AS M LEFT JOIN
                        (SELECT mastercode1, Mastercode2, sum(value1) As MTB, sum(value2) AS ATB From 
                        tran2  Where rectype = 2
                        And Date <=#" & Constant.CL_DATE & "# 
                        group by Mastercode1, Mastercode2) AS S1
                        On (S1.Mastercode1 = M.c) And (S1.Mastercode2 = M.CM)  ORDER BY M.Name"

        ElseIf qName = "GetProductInfo" Then
            RetQrr = "Select M.Name,M.Alias,M.PrintName,M.Code,M.D2,M.D3,M.D4,M.D9,M.D10,M.D16,M.D17,A.Address1,A.Address2,A.Address3,A.Address4 from Master1 AS M, MasterAddressInfo AS A where M.MasterType=6 AND M.Alias='" & ItemAlias & "' AND A.MasterCode=M.Code"
        ElseIf qName = "STPTName" Then
            RetQrr = "Select M1.Name,MS.* from Master1 as M1 inner join MasterSupport as MS on 
                        M1.Code=MS.MasterCode where MS.MasterCode=(Select CM8 From Master1 where 
                        MasterType=6 and 
                        Alias='" & ItemAlias & "')"
        ElseIf qName = "LastBill" Then
            RetQrr = "Select TOP 1 VchNo,VchAmtBaseCur,VchSalePurcAmt  from Tran1 ORDER BY VchNo DESC"
        ElseIf qName = "FindBill" Then
            RetQrr = "Select * from Tran1 where VchNo = '" & ItemAlias & "'"
        ElseIf qName = "MatCentre" Then
            RetQrr = "Select distinct MasterCode2,(select Name from Master1 where Code=Tran2.MasterCode2) as MCName,
                        (select ParentGrp from Master1 where Code=Tran2.MasterCode2) as PGCode,
                        (select Name from Master1 where Code=(select ParentGrp from Master1 where Code=Tran2.MasterCode2))
                        as PGName from Tran2 where rectype=2"
        ElseIf qName = "TestQrr" Then
            RetQrr = "Select * from Master1 where MasterType=2"
        End If
        If Constant.CURRENT_MODE = "DEV" Then
            MsgBox(RetQrr, Title:=qName)
        End If
        Return RetQrr

    End Function


    Public Function XMLItemDetail(ItemName, Qty, Price, Amt, STPTName, sessId)

        XMLStr2 = XMLStr2 & "<STPTName>" & STPTName & "</STPTName><MasterName1>Cash</MasterName1>"
        XMLStr1 = XMLStr1 & "<ItemDetail><ItemName>" & ItemName & "</ItemName><Qty>" & Qty & "</Qty><Price>" & Price & "</Price><Amt>" & Amt & "</Amt></ItemDetail>"

        Using addXML1 = System.IO.File.CreateText(opsPath & sessId & "_xml1.txt")
            addXML1.WriteLine(XMLStr1)
        End Using
        Using addXML2 = System.IO.File.CreateText(opsPath & sessId & "_xml2.txt")
            addXML2.WriteLine(XMLStr2)
        End Using
        Return True
    End Function

    Public Shared Function generateXML(VchDate, XMLStr1, XMLStr2)
        Dim XMLStr As String
        XMLStr = "<Sale>"
        XMLStr = XMLStr & "<VchSeriesName>Main</VchSeriesName><Date>" & VchDate & "</Date><VchType>9</VchType>"
        XMLStr = XMLStr & XMLStr2
        XMLStr = XMLStr & "<ItemEntries>"
        XMLStr = XMLStr & XMLStr1
        XMLStr = XMLStr & "</ItemEntries>"
        'XMLStr = XMLStr & "<BillSundries>"
        'XMLStr = XMLStr & "<BSDetail><SrNo>1</SrNo><BSName>Discount</BSName><PercentVal>14.5</PercentVal><Amt>238.55</Amt></BSDetail>"
        'XMLStr = XMLStr & "</BillSundries>"
        XMLStr = XMLStr & "</Sale>"


        Return XMLStr

    End Function



End Class
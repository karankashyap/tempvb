' Query for the description of item via bar code  (alias)
'   Table definition For all the tables consisting important data
' Information I want is product description, price, opening stock etc. 



'as per our discussion here are the points we are working on
'we need the following from your side 
'other products like elaj


'Check Power BI and IBM watson for dashboard designs 

Public Function saleVch()
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
    DataControl.generatePurchaseXML()
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
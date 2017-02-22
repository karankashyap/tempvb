Public Class Constant
    Inherits Control


    'CHANGE THE FOLLOWING VALUES
    Public Const PRG_PATH As String = "D:\Program Files (x86)\BusyWin" ' Path to where BUSY is installed
    Public Const DATA_PATH As String = "D:\Program Files (x86)\BusyWin\DATA\" ' 
    Public Const COMPANY_CODE As String = "Comp0003" ' Primary Company Code as shown in BUSY
    Public Const FY_DATE As String = "01-04-2017" ' FY Date to be used in the app
    Public Const CL_DATE As String = "04-01-2017" ' FY Date to be used in the app
    Public Const CURRENT_MODE As String = "DEBUG" ' DEV / LIVE / DEBUG
    Public Const INVOICE_DIR As String = "\invoice_pdf\" ' Company Code as shown in BUSY, NOT USED
    Public Const OPS_DIR As String = "\ops\" ' Company Code as shown in BUSY, NOT USED


    'DO NOT CHANGE THE FOLLOWING VALUES
    Public Const DEFAULT_DLL As String = "Busy2L16" ' Busy2L16 
    Public Const DEFAULT_L14_DLL As String = "Busy2L14" '  Busy2L14
    Public Const DEFAULT_CLASS As String = "CFixedInterface" ' CFixedInterface
    Public Const COMPANY_CODE_ALTER As String = "Comp0002" ' Company Code as shown in BUSY, NOT USED
    Public Const VCH_TYPE As Integer = 9 ' VCH TYPE for sale


    'ERR_CODES, DO NOT MODIFY
    Public Const ERR_SALE As String = "ERR: 3FS01 - Sale could not be recorded."
    Public Const ERR_DBREAD As String = "ERR: 1FD01 - Database could not be read."
    Public Const ERR_BUSYEXE As String = "ERR: 0FE01 - A Busy Function could not be executed."
    Public Const ERR_PDF As String = "ERR: 4FP01 - PDF Invoice could not be generated."
    Public Const ERR_OPNPDF As String = "ERR: 4FP02 - PDF Invoice could not be opened."
    Public Const ERR_PDFDIR As String = "ERR: 0FP01 - Invoice Directory could not be created."

End Class

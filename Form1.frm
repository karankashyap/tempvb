'VERSION 5.00
'Begin VB.Form Form1 
'Caption         =   "Form1"
'ClientHeight    =   3090
'ClientLeft      =   60
'ClientTop       =   450
'ClientWidth     =   4680
'LinkTopic       =   "Form1"
'ScaleHeight     =   3090
'ScaleWidth      =   4680
'StartUpPosition =   3  'Windows Default
'  Begin VB.CommandButton Command1 
'Caption         =   "Command1"
'Height          =   495
'Left            =   1440
'TabIndex        =   0
'Top             =   1080
'Width           =   1215
'End
'End
'Attribute VB_Name = "Form1"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'Option Explicit On
'

Private Sub Command1_Click()
    Dim FI As Object
    Dim Qry As String
    Dim rst As Recordset

    FI = CreateObject("Busy2L16.CFixedInterface")


    'Opens the BUSY database - SQL server
    'FI.OpenCSDB "C:\Busywin\", "Rachna", "sa", "busy", "BusyComp0002"
    'OpenCSDB function has following parameters
    '"C:\Busywin\" - BusyPath where Busy is installed
    '"Rachna" - SQL Server Name
    '"sa" - SQL Server UserName
    '"busy" - SQL Server Password
    '"BusyComp0002" - Busy Company Code

    'OR

    'For Access Mode - To Open BUSY Database
    FI.OpenDB("D:\Program Files (x86)\BusyWin", "D:\Program Files (x86)\BusyWin\DATA\", "Comp0001")
    'OpenDB function has following parameters
    '"C:\Busywin\" - BusyPath where Busy is installed
    '"D:\Busy Data\" - Data Path
    '"Comp0001" - Busy Company Code

    MsgBox("Database Connected")


    'Fetch Values from Database through Queries
    Qry = "Select * from Tran1 where VchType = 9"
    'This query will return list of all the Sale vouchers as VchType=9 belong to SALE voucher in Tran1 table
    
    Qry = "Select * from Master1 where MasterType=6"
    'This query will return list of all the Item Masters as MasterType=6 belongs to Item Master in Master1 table.

    rst = FI.GetRecordset(Qry)

    If rst.RecordCount > 0 Then
        rst.MoveFirst
        rst.MoveLast
        Do While Not rst.EOF
            MsgBox(Trim$(rst!Name.Value))
            rst.MoveNext
        Loop
    End If
    
    
    'Close Busy Company which is opened thru OpenCSDB/OpenDB function
    FI.CloseDB

    FI = Nothing

End Sub

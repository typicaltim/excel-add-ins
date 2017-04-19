Sub RegisterBalancerDEV()

    'Variable Declarations
        Dim vnull As Integer
    
    'Formatting the transaction data
    
        'Rename the Worksheet to "Transaction Data"
        ActiveSheet.Name = "Transaction_Data"
        
        'Create Table From Data
            'Treat variable "transactionTable" as a table (aka "list object")
            Dim transactionTable As ListObject
            'Create a table from Columns A - Q and store it in transactionTable
            Set transactionTable = ActiveSheet.ListObjects.Add(xlSrcRange, Columns("A:Q"), xlYes)
            'Rename the table for easy reference
            transactionTable.Name = "Transaction_Table"
            'Stylize the table so it is pretty :)
            transactionTable.TableStyle = "TableStyleMedium2"
            'Add a new column to hold data we will be making to consolidate the credit card types into one category.
            ActiveWorkbook.ActiveSheet.ListObjects("Transaction_Table").ListColumns.Add
                'Change the header for that new column.
                ActiveWorkbook.ActiveSheet.ListObjects("Transaction_Table").HeaderRowRange.Cells(1, 18).Value = "Transaction Type"
                
                'Consolidate the different credit card types into a single category "Credit", so that the pivot table is easier to read later
                For Each CellRangeRead In ActiveSheet.Range(ActiveSheet.Range("O2"), ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp))
                    'Compare O column cells to the following cases, when a match is found - write the transaction type to Column R
                    Select Case CellRangeRead
                    
                        Case "Checking"
                        CellRangeRead.Offset(0, 3).Value = "Check"
                        
                        Case "Corporate checking"
                        CellRangeRead.Offset(0, 3).Value = "Check"
                        
                        Case "Discover"
                        CellRangeRead.Offset(0, 3).Value = "Credit"
                        
                        Case "Visa"
                        CellRangeRead.Offset(0, 3).Value = "Credit"
                        
                        Case "MasterCard"
                        CellRangeRead.Offset(0, 3).Value = "Credit"
                        
                        Case "American Express"
                        CellRangeRead.Offset(0, 3).Value = "Credit"
                        
                        Case Else
                        CellRangeRead.Offset(0, 3).Value = "Please Contact I.T."
                        
                    End Select
                    
                    Next CellRangeRead
            
    'Create Pivot Table From "Transaction_Data"
    
        'Select the Transaction Data worksheet
        ActiveWorkbook.Sheets("Transaction_Data").Select
        'Select the A1 cell within that worksheet
        Range("A1").Select
        'Create a Pivot Table from the transactionTable Table within the Transaction Data worksheet.
        Set transactionPTable = ActiveWorkbook.Sheets("Transaction_Data").PivotTableWizard
        'Rename the worksheet that is created for the Pivot Table
        ActiveSheet.Name = "Summary_Page"
        'Rename the pivot table to "transactionPTable
        ActiveSheet.PivotTables(1).Name = "transactionPTable"
        'Make the pivot table look pretty :)
        ActiveSheet.PivotTables("transactionPTable").TableStyle2 = "PivotStyleMedium6"
        
        'Organize transaction data into the pivot table
        
            'Add Transaction Type to the row field
            Set objfield = transactionPTable.PivotFields("Transaction Type")
                With objfield
                    .Orientation = xlRowField
                End With
            
            'Add Transaction Reference Number to the row field
            Set objfield = transactionPTable.PivotFields("Transaction Reference Number")
                With objfield
                    .Orientation = xlRowField
                    'Enable multiple filters
                    .EnableMultiplePageItems = True
                    
                    'If an error is encountered, ignore it and keep going
                    '       Not sure if this is a respectable way to handle this sort of situation, but I'm doing it this way for now.
                    On Error Resume Next
                        'Remove blank transaction reference numbers
                        .PivotItems("(blank)").Visible = False
                    'Reset error handling, please complain if something bad happens
                    On Error GoTo 0
                    
                End With
            
            'Add Client User to the column field
            Set objfield = transactionPTable.PivotFields("Client User")
                With objfield
                    .Orientation = xlColumnField
                End With
            
            'Add Amount the the data field
            Set objfield = transactionPTable.PivotFields("Amount")
                With objfield
                    .Orientation = xlDataField
                    'Sum the total of this data
                    .Function = xlSum
                    'Format the values so that it matches US currency notation
                    .NumberFormat = "$ #,##0.00"
                End With
            
            'Add Applications to the filter field, and hide cancelled transactions
            Set objfield = transactionPTable.PivotFields("Applications")
                With objfield
                    .Orientation = xlPageField
                    .CurrentPage = "(All)"
                    'Enable multiple filters
                    .EnableMultiplePageItems = True
                    
                    'If an error is encountered, ignore it and keep going
                    '       Not sure if this is a respectable way to handle this sort of situation, but I'm doing it this way for now.
                    On Error Resume Next
                        'Remove voided/rejected/declined transactions
                        .PivotItems("Credit Card Authorization(Reject),Credit Card Settlement(Ignore)").Visible = False
                        .PivotItems("Credit Card Settlement(Ignore),Credit Card Authorization(Reject)").Visible = False
                    'Reset error handling, please complain if something bad happens
                    On Error GoTo 0
                    
                End With

    'Filter the pivot table by username to make it easier on users to identify the content they are concerned about

        ' \\ METHOD 1: USER ENTERS NAME PRIOR TO FILTERING \\
        'Set variable as string
        Dim myUserName As String
        'Ask user to input their username
        myUserName = Application.InputBox(Title:="User Name Verification", prompt:="Please Enter User Name", Type:=2)
        'For every Client User, check if it matches the input usename and hide it if it doesn't - that way only the input user's transactions are listed
        
        Dim arr() As Variant
        
        For Each pvtItem In ActiveSheet.PivotTables("transactionPTable").PivotFields("Client User").PivotItems
            arr(UBound(arr)) = pvtItem.Name
            ReDim Preserve arr(1 To UBound(arr) + 1) As Variant
        Next pvtItem
        
            If IsInArray(myUserName, arr) Then
                For Each pvtItem In ActiveSheet.PivotTables("transactionPTable").PivotFields("Client User").PivotItems
                    pvtItem.Visible = False
                Next pvtItem
                ActiveSheet.PivotTables("transactionPTable").PivotFields("Client User").PivotItems.pvtItem(myUserName).Visible = True
            Else
                MsgBox (myUserName + " does not exist in this pivot table. Please enter a valid user name.")
            End If
        
        
        ' \\ METHOD 2: SYSTEM VARIABLE IS CALLED FOR AUTOMATIC FILTERING \\
        '           This sh*t doesn't f**king work. Don't bother.

        '        'Set variable as string
        '        Dim myUserName As String
        '        'Use environment variable to determine logged in user
        '        myUserName = Environ$("UserName")
        '        myUserName = (StrConv(Var, VbStrConv.vbLowerCase))
        
        '        'For every Client User, check if it matches the logged on user and hide it if it doesn't - that way only the current user's transactions are listed
        '            For Each pvtItem In ActiveSheet.PivotTables("transactionPTable").PivotFields("Client User").PivotItems
        '            If pvtItem.Name = myUserName Then
        '                    pvtItem.Visible = True
        '                Else
        '                    pvtItem.Visible = False
        '            End If
        '            Next pvtItem

End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

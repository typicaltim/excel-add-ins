Sub RegisterBalancer()

    'Variable Declarations
    
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
                        
                        Case "Visa"
                        CellRangeRead.Offset(0, 3).Value = "Credit"
                        
                        Case "MasterCard"
                        CellRangeRead.Offset(0, 3).Value = "Credit"
                        
                        Case "American Express"
                        CellRangeRead.Offset(0, 3).Value = "Credit"
                        
                        Case "Discover"
                        CellRangeRead.Offset(0, 3).Value = "Check"
                        
                        Case Else
                        CellRangeRead.Offset(0, 3).Value = Null
                        
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
            Set objField = transactionPTable.PivotFields("Transaction Type")
            objField.Orientation = xlRowField
            
            'Add Transaction Reference Number to the row field
            Set objField = transactionPTable.PivotFields("Transaction Reference Number")
            objField.Orientation = xlRowField
            
            'Add Client User to the column field
            Set objField = transactionPTable.PivotFields("Client User")
            objField.Orientation = xlColumnField
            
            'Add Amount the the data field
            Set objField = transactionPTable.PivotFields("Amount")
            objField.Orientation = xlDataField
                'Sum the total of this data
                objField.Function = xlSum
                'Format the values so that it matches US currency notation
                objField.NumberFormat = "$ #,##0.00"
    
    'Filter the pivot table to make it easier on users to identify the content they are concerned about
        
        'Set variable as string
        Dim myUserName As String
        'Use environment variable to determine logged in user
        myUserName = Environ$("UserName")
        
        'For every Client User, check if it matches the logged on user and hide it if it doesn't - that way only the current user's transactions are listed
        For Each pvtItem In ActiveSheet.PivotTables("transactionPTable").PivotFields("Client User").PivotItems
            If pvtItem.Name = myUserName Then
                pvtItem.Visible = True
            Else
                pvtItem.Visible = False
            End If
            Next pvtItem

End Sub

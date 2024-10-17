# Automated Financial Accounting System: From Transaction Entry to Financial Statements in Excel
## Objective
The objective of this project is to develop an automated financial accounting system in Excel that streamlines the recording and processing of transactions. The system will automatically post entries to the General Ledger, extract the Trial Balance, and generate key financial statements, including the Profit & Loss Account and Balance Sheet. This solution aims to enhance efficiency, reduce manual errors, and provide accurate, real-time financial reporting for users.
## Link to the Excel project
- <a href="https://github.com/Gideon-Mensah/fmexcelaccount/blob/main/Financial%20Account%20Modeling%20github.xlsm">FMmodel</a>
## How to used the project
- Download the project file from the above link
- Upon opening the Excel file, you will be directed to the Dashboard sheet automatically
  <img src="https://github.com/Gideon-Mensah/fmexcelaccount/blob/main/Dashboard.jpg"> 
- Click on the 'Data' button to navigate to a sheet where you can set up your business name and account structure. The chart of accounts has been pre-configured and should not be modified, as changes could impact the generation of the financial statements. However, you can create new accounts and link them to the pre-determined chart of accounts. You can click on the dashboard button to return to dashboard.
  <img src="https://github.com/Gideon-Mensah/fmexcelaccount/blob/main/Data.jpg">
- On the Dashboard, you can navigate between financial statements by clicking the available buttons. To record transactions, simply enter the details in the specified textboxes under the 'Particulars' section, input the account to be debited under the 'Debit' section and the account to be credited under the 'Credit' section, and then click 'Submit.' The system will automatically post the entry to the General Ledger, update the Trial Balance, generate a Profit & Loss statement if needed, and create the Balance Sheet.
- General Ledger
  <img src="https://github.com/Gideon-Mensah/fmexcelaccount/blob/main/General%20Ledger.jpg">
- Trial Balance
  <img src="https://github.com/Gideon-Mensah/fmexcelaccount/blob/main/Trial%20Balance.jpg">
- Income Statement
  <img src="https://github.com/Gideon-Mensah/fmexcelaccount/blob/main/Income%20Statement.jpg">
- Balance Sheet
  <img src="https://github.com/Gideon-Mensah/fmexcelaccount/blob/main/Balance%20Sheet.jpg">
## Code for this project
This model utilizes VBA, so you'll need to enable the Developer tab in Excel to create your own version of this project. Provided below is the VBA code for this project
- ### Code for recording transactions and generating the various financial statements

      Dim iRow As Long
      Dim ws As Worksheet
      Set ws = Worksheets("GeneralLedger")
        iRow = ws.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1
        ws.Cells(iRow, 2).Value = transactiondate.Value
        ws.Cells(iRow, 3).Value = Description.Value
        ws.Cells(iRow, 4).Value = PVnumber.Value
        ws.Cells(iRow, 5).Value = accountdebitname.Value
        ws.Cells(iRow, 6).Value = chartofaccountdebit.Value
        ws.Cells(iRow, 7).Value = amountdebitamount.Value
        iRow = ws.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1
        ws.Cells(iRow, 2).Value = transactiondate.Value
        ws.Cells(iRow, 3).Value = Description.Value
        ws.Cells(iRow, 4).Value = PVnumber.Value
        ws.Cells(iRow, 5).Value = accountnamecredit.Value
        ws.Cells(iRow, 6).Value = chartofaccountcredit.Value
        ws.Cells(iRow, 8).Value = amountcredit.Value
        CreateTrialBalanceInExistingSheet
        CreateProfitAndLossFromTrialBalance
        CreateVerticalBalanceSheet
        MsgBox "Data saved successfully!", vbInformation

      Sub CreateTrialBalanceInExistingSheet()
        Dim wsData As Worksheet
        Dim wsTrialBalance As Worksheet
        Dim lastRow As Long
        Dim trialBalanceDict As Object
        Dim customSortOrder As Object ' Declare customSortOrder as a local variable
        Dim i As Long
        Dim chartOfAccount As String
        Dim accountName As String
        Dim debitAmount As Double
        Dim creditAmount As Double
        Dim consolidatedData As Variant
        Dim rowOffset As Long
        Dim totalDebit As Double
        Dim totalCredit As Double
        Dim key As Variant
        Set trialBalanceDict = CreateObject("Scripting.Dictionary")
        Set customSortOrder = CreateObject("Scripting.Dictionary")
        customSortOrder.Add "Fixed Assets", 1
        customSortOrder.Add "Current Assets", 2
        customSortOrder.Add "Current Liabilities", 3
        customSortOrder.Add "Long-term Liabilities", 4
        customSortOrder.Add "Capital", 5
        customSortOrder.Add "Revenue", 6
        customSortOrder.Add "Cost of Sales", 7
        customSortOrder.Add "General Administration", 8
        ' Set the worksheet where the data is stored (adjust sheet name if necessary)
        Set wsData = ThisWorkbook.Sheets("GeneralLedger") ' Change Data to your actual data sheet name
        ' Get the last row of data in Column A (account names)
        lastRow = wsData.Cells(wsData.Rows.Count, "B").End(xlUp).Row
        ' Loop through the data to consolidate accounts by Chart of Accounts
        For i = 5 To lastRow ' Assuming the data starts in row 2
        accountName = wsData.Cells(i, 5).Value
        chartOfAccount = wsData.Cells(i, 6).Value ' Assuming Chart of Accounts is in column D (4th column)
        debitAmount = wsData.Cells(i, 7).Value
        creditAmount = wsData.Cells(i, 8).Value
        ' Create a unique key combining Chart of Account and Account Name
        ' Dim key As Variant
        key = chartOfAccount & " - " & accountName
        ' Check if the combined key already exists in the dictionary
        If trialBalanceDict.Exists(key) Then
            ' Add debit and credit to the existing account
            consolidatedData = trialBalanceDict(key)
            consolidatedData(0) = consolidatedData(0) + debitAmount
            consolidatedData(1) = consolidatedData(1) + creditAmount
            trialBalanceDict(key) = consolidatedData
        Else
            ' Add new account to the dictionary
            trialBalanceDict.Add key, Array(debitAmount, creditAmount)
        End If
        Next i
        ' Set the worksheet where the trial balance will be inserted (adjust sheet name if necessary)
        Set wsTrialBalance = ThisWorkbook.Sheets("TrialBalance") ' Adjust the name to the sheet you want to insert
        ' Starting point to insert trial balance data
        Dim startRow As Long
        Dim startCol As Long
        startRow = 3 ' Change this to the row where you want to start output
        startCol = 2 ' Change this to the column where you want the result (E = 5)
        ' Add headers
        wsTrialBalance.Cells(startRow, startCol).Value = "Chart of Account"
        wsTrialBalance.Cells(startRow, startCol + 1).Value = "Account Name"
        wsTrialBalance.Cells(startRow, startCol + 2).Value = "Debit"
        wsTrialBalance.Cells(startRow, startCol + 3).Value = "Credit"  
        ' Sort the keys based on custom order
        Dim sortedKeys() As String
        ReDim sortedKeys(1 To trialBalanceDict.Count) 
        ' Add keys to the array
        Dim k As Long
        k = 1
        For Each key In trialBalanceDict.Keys
        sortedKeys(k) = key
        k = k + 1
        Next key
       ' Manually sort the array using the custom sort order, passing customSortOrder to the subroutine
        Call SortKeysByCustomOrder(sortedKeys, customSortOrder)  
      ' Initialize total debit and credit
        totalDebit = 0
        totalCredit = 0  
        ' Write consolidated and sorted data to the existing sheet
        rowOffset = startRow + 1 ' Start from the row after headers
        For k = 1 To UBound(sortedKeys)
        key = sortedKeys(k)
        consolidatedData = trialBalanceDict(key)     
        ' Split the combined key to separate Chart of Account and Account Name
        Dim parts() As String
        parts = Split(key, " - ")
        ' Insert data into sheet
        wsTrialBalance.Cells(rowOffset, startCol).Value = parts(0) ' Chart of Account
        wsTrialBalance.Cells(rowOffset, startCol + 1).Value = parts(1) ' Account Name
        wsTrialBalance.Cells(rowOffset, startCol + 2).Value = consolidatedData(0) ' Debit
        wsTrialBalance.Cells(rowOffset, startCol + 3).Value = consolidatedData(1) ' Credit  
        ' Add to totals
        totalDebit = totalDebit + consolidatedData(0)
        totalCredit = totalCredit + consolidatedData(1)
        rowOffset = rowOffset + 1
        Next k
        ' Add totals at the end
        wsTrialBalance.Cells(rowOffset, startCol).Value = "Total"
        wsTrialBalance.Cells(rowOffset, startCol + 2).Value = totalDebit
        wsTrialBalance.Cells(rowOffset, startCol + 3).Value = totalCredit
        ' Format the total row (optional)
        wsTrialBalance.Cells(rowOffset, startCol).Font.Bold = True
        wsTrialBalance.Cells(rowOffset, startCol + 2).Font.Bold = True
        wsTrialBalance.Cells(rowOffset, startCol + 3).Font.Bold = True
        End Sub
        Sub SortKeysByCustomOrder(arr() As String, customSortOrder As Object)
        Dim i As Long, j As Long
        Dim temp As String
        Dim key1 As String, key2 As String
        Dim parts1() As String, parts2() As String
        Dim value1 As Integer, value2 As Integer
        ' Perform bubble sort based on custom order
        For i = LBound(arr) To UBound(arr) - 1
            For j = i + 1 To UBound(arr)
                parts1 = Split(arr(i), " - ")
                parts2 = Split(arr(j), " - ")
                ' Determine the sort order values
                If customSortOrder.Exists(parts1(0)) Then
                    value1 = customSortOrder(parts1(0))
                Else
                    value1 = 9999 ' Assign a large number if not found in sort order
                End If
                If customSortOrder.Exists(parts2(0)) Then
                    value2 = customSortOrder(parts2(0))
                Else
                    value2 = 9999 ' Assign a large number if not found in sort order
                End If
                ' Swap if out of order
                If value1 > value2 Then
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                End If
              Next j
          Next i
        End Sub
        Sub CreateProfitAndLossFromTrialBalance()
        Dim wsTrialBalance As Worksheet
        Dim wsPL As Worksheet
        Dim lastRow As Long
        Dim revenueTotal As Double
        Dim costOfSalesTotal As Double
        Dim expensesTotal As Double
        Dim grossProfit As Double
        Dim netProfit As Double
        Dim rowOffset As Long
        Dim chartOfAccount As String
        Dim accountName As String
        Dim debitAmount As Double
        Dim creditAmount As Double
        Dim accountType As String
        ' Set the worksheet containing the trial balance
        Set wsTrialBalance = ThisWorkbook.Sheets("TrialBalance") ' Change this to your trial balance sheet name
        ' Set the existing worksheet where the P&L will be inserted
        Set wsPL = ThisWorkbook.Sheets("Profit or Loss ac") ' Adjust this to your P&L sheet name
        ' Clear previous P&L data if necessary
        ' wsPL.Cells.Clear
        wsPL.Rows("5:" & wsPL.Rows.Count).ClearContents
        ' Initialize totals
        revenueTotal = 0
        costOfSalesTotal = 0
        expensesTotal = 0
        ' Get the last row of the trial balance
        lastRow = wsTrialBalance.Cells(wsTrialBalance.Rows.Count, "B").End(xlUp).Row
        ' Define starting row for the P&L statement in the existing sheet
        rowOffset = 4
        ' Write headers for the P&L
        ' wsPL.Cells(1, 1).Value = "Profit and Loss Account"
        ' wsPL.Cells(rowOffset, 1).Value = "Description"
        ' wsPL.Cells(rowOffset, 2).Value = "Amount"
        rowOffset = rowOffset + 1
        ' Loop through the trial balance and categorize accounts
        For i = 4 To lastRow ' Assuming data starts in row 2
        chartOfAccount = wsTrialBalance.Cells(i, 2).Value ' Adjust for the correct column containing Chart of Account
        accountName = wsTrialBalance.Cells(i, 3).Value ' Adjust for the correct column containing account names
        debitAmount = wsTrialBalance.Cells(i, 4).Value ' Assuming debit column is column C
        creditAmount = wsTrialBalance.Cells(i, 5).Value ' Assuming credit column is column D
        ' Check if the account belongs to Revenue, Cost of Sales, or Expenses
        accountType = GetAccountTypeByChart(chartOfAccount) 
        Select Case accountType
            Case "Revenue"
                wsPL.Cells(rowOffset, 1).Value = "Revenue:"
                rowOffset = rowOffset + 1
                wsPL.Cells(rowOffset, 1).Value = accountName
                wsPL.Cells(rowOffset, 2).Value = creditAmount
                revenueTotal = revenueTotal + creditAmount
                rowOffset = rowOffset + 1      
                wsPL.Cells(rowOffset, 1).Font.Bold = False
                wsPL.Cells(rowOffset, 2).Font.Bold = False 
            Case "Cost of Sales"
                wsPL.Cells(rowOffset, 1).Value = "Cost of Sales:"
                rowOffset = rowOffset + 1
                wsPL.Cells(rowOffset, 1).Value = accountName
                wsPL.Cells(rowOffset, 2).Value = debitAmount
                costOfSalesTotal = costOfSalesTotal + debitAmount
                rowOffset = rowOffset + 1    
                wsPL.Cells(rowOffset, 1).Font.Bold = False
                wsPL.Cells(rowOffset, 2).Font.Bold = False   
                ' Calculate Gross Profit
                grossProfit = revenueTotal - costOfSalesTotal
                ' Write Gross Profit
                rowOffset = rowOffset + 1
                wsPL.Cells(rowOffset, 1).Value = "Gross Profit"
                wsPL.Cells(rowOffset, 2).Value = grossProfit
                wsPL.Cells(rowOffset, 1).Font.Bold = True
                wsPL.Cells(rowOffset, 2).Font.Bold = True
                rowOffset = rowOffset + 1
              Case "Expenses"
                wsPL.Cells(rowOffset, 1).Value = "Expenses:"
                rowOffset = rowOffset + 1
                wsPL.Cells(rowOffset, 1).Value = accountName
                wsPL.Cells(rowOffset, 2).Value = debitAmount
                expensesTotal = expensesTotal + debitAmount
                rowOffset = rowOffset + 1 
                wsPL.Cells(rowOffset, 1).Font.Bold = False
                wsPL.Cells(rowOffset, 2).Font.Bold = False  
                ' Calculate Net Profit
                netProfit = revenueTotal - expensesTotal - costOfSalesTotal
                ' Write Net Profit
                rowOffset = rowOffset + 1
                wsPL.Cells(rowOffset, 1).Value = "Net Profit"
                wsPL.Cells(rowOffset, 2).Value = netProfit
                wsPL.Cells(rowOffset, 1).Font.Bold = True
                wsPL.Cells(rowOffset, 2).Font.Bold = True    
            End Select
        Next i
        'MsgBox "Profit and Loss Account created successfully!", vbInformation
        End Sub
        ' Function to categorize accounts by Chart of Account (Revenue, Cost of Sales, Expenses)
        Function GetAccountTypeByChart(chartOfAccount As String) As String
        ' Example of how to categorize based on chart of account ranges or patterns
        Select Case chartOfAccount
            Case "Revenue"
                GetAccountTypeByChart = "Revenue"
            Case "Cost of Sales"
                GetAccountTypeByChart = "Cost of Sales"
            Case "General Administration"
                GetAccountTypeByChart = "Expenses"
        End Select
        End Function
        Sub CreateVerticalBalanceSheet()
        Dim wsTrial As Worksheet
        Dim wsBalanceSheet As Worksheet
        Dim lastRow As Long
        Dim rowBalanceSheet As Long
        Dim totalfixedAssets As Double
        Dim totalcurrentAssets As Double
        Dim totalAssets As Double
        Dim totallongtermLiabilities As Double
        Dim totalcurrentLiabilities As Double
        Dim totalLiabilities As Double
        Dim totalEquity As Double
        Dim totalRevenue As Double
        Dim totalcostofsalesExpenses As Double
        Dim totalgeneraladminExpenses As Double
        Dim netProfit As Double
        ' Set references to sheets
        Set wsTrial = ThisWorkbook.Sheets("TrialBalance") ' Change to your Trial Balance sheet name
        Set wsBalanceSheet = ThisWorkbook.Sheets("Balance Sheet") ' Change to the existing Balance Sheet sheet name
        ' Clear the existing Balance Sheet contents before starting
        wsBalanceSheet.Rows("3:" & wsBalanceSheet.Rows.Count).ClearContents
        ' Add heading for the Balance Sheet
        wsBalanceSheet.Cells(3, 1).Value = "Balance Sheet"
        wsBalanceSheet.Cells(3, 2).Value = "As of: " & Date
        rowBalanceSheet = 4 ' Start from row 4 on the Balance Sheet
        ' Add Assets heading
        wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Assets"
        wsBalanceSheet.Cells(rowBalanceSheet, 1).Font.Bold = True
        rowBalanceSheet = rowBalanceSheet + 1
        ' Add Fixed Assets heading
        wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Fixed Assets"
        wsBalanceSheet.Cells(rowBalanceSheet, 1).Font.Bold = True
        rowBalanceSheet = rowBalanceSheet + 1
        ' Find the last row of the trial balance
        lastRow = wsTrial.Cells(wsTrial.Rows.Count, 2).End(xlUp).Row
        ' Loop through the trial balance and populate the balance sheet
        Dim i As Long
        For i = 4 To lastRow ' Assuming row 1 is headers
        Dim accountName As String
        Dim debitAmount As Double
        Dim creditAmount As Double
        Dim accountType As String
        accountType = wsTrial.Cells(i, 2).Value ' Account name in column B
        accountName = wsTrial.Cells(i, 3).Value ' Debit amount in column C
        debitAmount = wsTrial.Cells(i, 4).Value ' Credit amount in column D
        creditAmount = wsTrial.Cells(i, 5).Value ' Account type in column E
        Select Case accountType
            ' Fixed Assets
            Case "Fixed Assets"
                wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = accountName
                wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = debitAmount - creditAmount
                totalfixedAssets = totalfixedAssets + (debitAmount - creditAmount)
                rowBalanceSheet = rowBalanceSheet + 1
            wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Total Fixed Assets"
            wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = totalfixedAssets
            rowBalanceSheet = rowBalanceSheet + 1 
            ' Current Assets
            Case "Current Assets"
                If wsBalanceSheet.Cells(rowBalanceSheet - 1, 1).Value <> "Current Assets" Then
                    wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Current Assets"
                    wsBalanceSheet.Cells(rowBalanceSheet, 1).Font.Bold = True
                    rowBalanceSheet = rowBalanceSheet + 1
                End If
                wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = accountName
                wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = debitAmount - creditAmount
                totalcurrentAssets = totalcurrentAssets + (debitAmount - creditAmount)
                rowBalanceSheet = rowBalanceSheet + 1
                wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Total Current Assets"
                wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = totalcurrentAssets
                rowBalanceSheet = rowBalanceSheet + 1
                wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Total Assets"
                wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = totalfixedAssets + totalcurrentAssets
                rowBalanceSheet = rowBalanceSheet + 2 ' Leave space
            ' Long-Term Liabilities
              Case "Long term Liabilities"
                If wsBalanceSheet.Cells(rowBalanceSheet - 1, 1).Value <> "Long term Liabilities" Then
                    wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Long term Liabilities"
                    wsBalanceSheet.Cells(rowBalanceSheet, 1).Font.Bold = True
                    rowBalanceSheet = rowBalanceSheet + 1
                End If
                wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = accountName
                wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = creditAmount - debitAmount
                totallongtermLiabilities = totallongtermLiabilities + (creditAmount - debitAmount)
                rowBalanceSheet = rowBalanceSheet + 1
                wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Total Long-Term Liabilities"
                wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = totallongtermLiabilities
                rowBalanceSheet = rowBalanceSheet + 1
            ' Current Liabilities
              Case "Current Liabilities"
                If wsBalanceSheet.Cells(rowBalanceSheet - 1, 1).Value <> "Current Liabilities" Then
                    wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Current Liabilities"
                    wsBalanceSheet.Cells(rowBalanceSheet, 1).Font.Bold = True
                    rowBalanceSheet = rowBalanceSheet + 1
                End If
                wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = accountName
                wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = creditAmount - debitAmount
                totalcurrentLiabilities = totalcurrentLiabilities + (creditAmount - debitAmount)
                rowBalanceSheet = rowBalanceSheet + 1
                wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Total Current Liabilities"
                wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = totalcurrentLiabilities
                rowBalanceSheet = rowBalanceSheet + 1  
                wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Total Liabilities"
                wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = totallongtermLiabilities + totalcurrentLiabilities
                rowBalanceSheet = rowBalanceSheet + 2 ' Leave space
            ' Equity
              Case "Capital"
                If wsBalanceSheet.Cells(rowBalanceSheet - 1, 1).Value <> "Capital" Then
                    rowBalanceSheet = rowBalanceSheet + 1
                    wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Capital"
                    wsBalanceSheet.Cells(rowBalanceSheet, 1).Font.Bold = True
                    rowBalanceSheet = rowBalanceSheet + 1
                End If
                wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = accountName
                wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = creditAmount - debitAmount
                totalEquity = totalEquity + (creditAmount - debitAmount)
                rowBalanceSheet = rowBalanceSheet + 1
            ' Revenue and Expense for Net Profit calculation
            Case "Revenue"
                totalRevenue = totalRevenue + (creditAmount - debitAmount)
            Case "Cost of Sales"
                totalcostofsalesExpenses = totalcostofsalesExpenses + (debitAmount - creditAmount)
            Case "General Administration"
                totalgeneraladminExpenses = totalgeneraladminExpenses + (debitAmount - creditAmount)
            End Select
        Next i
        ' Calculate Net Profit (Revenue - Expenses)
        netProfit = totalRevenue - totalgeneraladminExpenses - totalcostofsalesExpenses
        ' Add Net Profit to the Equity section
        totalEquity = totalEquity + netProfit
        ' Display Net Profit under Equity
        wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Net Profit"
        wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = netProfit
         ' Move to the next row to add totals and labels
        rowBalanceSheet = rowBalanceSheet + 1
        wsBalanceSheet.Cells(rowBalanceSheet, 1).Value = "Total Liabilities & Equity"
        wsBalanceSheet.Cells(rowBalanceSheet, 2).Value = totalEquity + totallongtermLiabilities + totalcurrentLiabilities
      End Sub
  
  - ### Code for navigating in and out of the various sheets (NB: Sample for one navigation. the excel file contain all the codes)
        Dim sheetName As String
        sheetName = "Profit or Loss ac" 'Replace "Sheet1" with the exact name of your sheet
        ' Check if the sheet exists before trying to unhide it
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = Sheets(sheetName)
        On Error GoTo 0
        If Not ws Is Nothing Then
            ' If the sheet exists, make it visible
            ws.Visible = True
            ws.Activate
        Else
            MsgBox "Sheet '" & sheetName & "' does not exist.", vbExclamation, "Error"
        End If

    - ### Code for disabling worksheet tab and fomular bar
           Application.DisplayFormulaBar = False
           ActiveWindow.DisplayWorkbookTabs = False
      
    ## Conculsion

    This Excel-based financial accounting system, incorporating automated processes for transaction entry and financial statement generation, is designed to enhance efficiency and accuracy in managing financial data. With the use of VBA, the model streamlines the workflow from recording transactions to producing the General Ledger, Trial Balance, Profit & Loss Account, and Balance Sheet. The full project, including the VBA code and documentation, has been made publicly available on GitHub, allowing others to review, use, and further develop the model for their own needs.




    

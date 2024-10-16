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


    

Sub classifyData()
    
    '重要!! 要修改的變數
    ' customerSheetName 待整理的工作簿名稱
    ' lastSelectColum 表頭的最後一欄
    '
    customerSheetName = "客戶明細"
    lastSelectColum = "K"

    '**********************************
    
    Dim Name As String
    Dim findedName As String
    Dim checkSheet As Worksheet
    
    customerSheet_RowIndex = 1
    customerSheet_FirstCompanyNameIndex = 2

    '取得excel中最大的列數
    totalRow = Sheets(customerSheetName).Rows.Count

    '從最後一列往上，找到第一個有值的儲存格，並回傳該列數
    finalRow = Sheets(customerSheetName).Cells(totalRow, 2).End(xlUp).Row

    Do
        customerSheet_RowIndex = customerSheet_RowIndex + 1

        '在客戶明細的工作簿中，取得 B 欄的公司名稱
        Name = Sheets(customerSheetName).Range("B" & customerSheet_RowIndex).Value
   
        '如果沒有對應公司名稱的工作簿
        If checkSheetName(Name) = False Then
            '就建一個新的，並把客戶明細工作簿的表頭複製到新建的工作簿中
            Sheets.Add.Name = Name
            Sheets(customerSheetName).Activate
            Sheets(customerSheetName).Range("A1:" & lastSelectColum & "1").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets(Name).Activate
            Sheets(Name).Select
            ActiveSheet.Paste
        End If

        If findedName = "" Then
            findedName = Name
        End If


        '如果:
        '   1.查詢的公司名稱與客戶明細工作簿中目前選取的名稱 不同 且 未到 最後一筆資料
        '   或
        '   2.查詢的公司名稱與客戶明細工作簿中目前選取的名稱 相同 且 到 最後一筆資料
        
        '就將找到的資料貼到對應的工作簿中
        '並調整 整理完畢的工作簿 欄位大小
        If (findedName <> Name And customerSheet_RowIndex <> finalRow) Or (findedName = Name And customerSheet_RowIndex = finalRow) Then

            '如果到了最後一筆資料，選取剩下的資料
            '不是的話就選取目前列數的前一筆，因為目前列數的資料是下一家公司的資料
            If customerSheet_RowIndex = finalRow Then
                RowNumber = customerSheet_RowIndex
            Else
                RowNumber = customerSheet_RowIndex - 1
            End If

            '將之前 剪下 或 複製 的模式去除
            '已確保要 剪下 或 複製 的是現在選取到的資料
            Application.CutCopyMode = False
    
            '將此公司的資料複製，貼到對應的工作簿
            Sheets(customerSheetName).Range("A" & customerSheet_FirstCompanyNameIndex & ":" & lastSelectColum & RowNumber).Copy
            Sheets(findedName).Activate
            Sheets(findedName).Range("A2").Select
            ActiveSheet.Paste
        
            jusitfyCol (findedName)
            
            customerSheet_FirstCompanyNameIndex = customerSheet_RowIndex
        End If
        
        findedName = Sheets(customerSheetName).Range("B" & customerSheet_RowIndex).Value

    '若還未處理到客戶明細中的最後一列資料，就繼續執行迴圈
    Loop While customerSheet_RowIndex < finalRow
    
End Sub

'檢查活頁是否存在
Function checkSheetName(sheetName As String)
        isfind = False
        For Each st In Sheets
            If st.Name = sheetName Then
               isfind = True
               Exit For
            End If
        Next
        checkSheetName = isfind
End Function

'調整欄位大小
Function jusitfyCol(sheetName As String)
    Sheets(sheetName).Activate

    '取得excel中最大的欄數
    totalColumn = Sheets(sheetName).Columns.Count

    '從最後一欄往左，找到第一個有值的儲存格，並回傳該欄數
    finalColumn = Sheets(sheetName).Cells(1, totalColumn).End(xlToLeft).Column
    
    '從第一欄開始做大小的調整
    For i = 1 To finalColumn
        Sheets(sheetName).Range(Columns(i), Columns(i)).EntireColumn.AutoFit
    Next i
End Function


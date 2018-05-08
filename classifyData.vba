Sub classifyData()
    
    Dim Name As String
    Dim findedName As String
    Dim checkSheet As Worksheet
    
    customerSheetRowIndex = 1
    companySheetRowIndex = 1

    '取得excel中最大的列數
    totalRow = Sheets("客戶明細").Rows.Count

    '從最後一列往上，找到第一個有值的儲存格，並回傳該列數
    finalRow = Sheets("客戶明細").Cells(totalRow, 2).End(xlUp).Row

    Do
        customerSheetRowIndex = customerSheetRowIndex + 1
        companySheetRowIndex = companySheetRowIndex + 1

        '在客戶明細的工作簿中，取得 B 欄的公司名稱
        Name = Sheets("客戶明細").Range("B" & customerSheetRowIndex).Value
   
        '如果沒有對應公司名稱的工作簿
        If checkSheetName(Name) = False Then
            '就建一個新的，並把客戶明細工作簿的表頭複製到新建的工作簿中
            Sheets.Add.Name = Name
            Sheets("客戶明細").Activate
            Sheets("客戶明細").Range("A1:K1").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets(Name).Activate
            Sheets(Name).Select
            ActiveSheet.Paste
        End If

        If findedName = "" Then
            findedName = Name
        End If

        '如果查詢的公司名稱與客戶明細工作簿中目前選取的名稱不同
        '就將行數初始化，從工作簿的第2列開始
        '並調整 整理完畢的工作簿 欄位大小
        If findedName <> Name Then
            companySheetRowIndex = 2
            jusitfyCol (findedName)
        End If
        
        Set findcell = Sheets("客戶明細").Range("b1:b" & finalRow).Find(what:=Name, LookIn:=xlValues)
        
        findedName = findcell.Value

        '將之前 剪下 或 複製 的模式去除
        '已確保要 剪下 或 複製 的是現在選取到的資料
        Application.CutCopyMode = False

        '將 A 欄 到 K 欄 的資料剪下
        Sheets("客戶明細").Range("A" & customerSheetRowIndex & ":K" & customerSheetRowIndex).Cut
        Sheets(Name).Activate
        Sheets(Name).Range("A" & companySheetRowIndex).Select
        ActiveSheet.Paste

    '若還未處理到客戶明細中的最後一列資料，就繼續執行迴圈
    Loop While customerSheetRowIndex < finalRow
    
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

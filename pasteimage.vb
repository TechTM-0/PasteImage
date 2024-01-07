Sub InsertImagesToExcel()
    ' フォルダのパスを指定（初期値: xxxx）
    Dim folderPath As String
    folderPath = "取得したい画像フォルダ"
    
    ' 宣言
    Dim excelApp As Excel.Application
    Dim Workbook As Excel.Workbook
    Dim Worksheet As Excel.Worksheet
    Dim imageCell As Excel.Range
    Dim imageWidth As Double
    Dim imageHeight As Double
    Dim Over_Height As Boolean '//高さが409より大きいときにTRUE
    Dim h As Double '//高さに対する幅の比
    
    ' Excelアプリケーションを作成
    Set excelApp = New Excel.Application
    
    ' Excelブックを追加
    Set Workbook = excelApp.Workbooks.Add
    
    ' シート1を選択
    Set Worksheet = Workbook.Sheets(1)
    
    ' 画像の挿入位置を指定（A1）
    Set imageCell = Worksheet.Range("A1")
    
    ' フォルダ内の画像を取得して、Excelに貼り付け
    Dim imageFile As String
    imageFile = Dir(folderPath & "\*.JPG") ' 拡張子が.JPGの画像のみを対象にする場合
    

    Do While imageFile <> ""
        ' 画像を挿入

        With Worksheet.Pictures.Insert(folderPath & "\" & imageFile)

            ' 画像の大きさを取得
            imageWidth = .Width
            imageHeight = .Height
            
            '画像の高さが409pointを超えているかチェック
            Over_Height = Check_IMGHEIGHT(imageHeight)
            
            '-----------------------------------------------------------------
            '画像サイズをアスペクト比を維持し縮小
            If Over_Height Then
                '比を計算
                h = .Width / .Height
                
                imageHeight = 407
                imageWidth = .Height * h  
            End If
            '-----------------------------------------------------------------
            
            ' 画像をセル内に収めるために行と列のサイズを調整
            imageCell.RowHeight = imageHeight + 2
            imageCell.ColumnWidth = imageWidth / 7
            
            '位置
            .ShapeRange.LockAspectRatio = msoFalse
            .Height = imageHeight
            .Width = imageWidth
            .Top = imageCell.Top + 1
            .Left = imageCell.Left + 1
            
            ' 次の画像を挿入するセルを指定（A列の次の行）
            Set imageCell = Worksheet.Cells(imageCell.Row + 1, imageCell.Column)
        End With
        
        ' 次の画像を探す
        imageFile = Dir
    Loop
    
    ' Excelアプリケーションを表示
    excelApp.Visible = True
    
    ' メモリ解放
    Set imageCell = Nothing
    Set Worksheet = Nothing
    Set Workbook = Nothing
    Set excelApp = Nothing
End Sub

Function Check_IMGHEIGHT(v) As Boolean
'//true:409より大きい
'//false:409以内

    Check_IMGHEIGHT = False
    If v > 409 Then Check_IMGHEIGHT = True
End Function


Attribute VB_Name = "CopyRecordbyRecord"
Option Compare Database

Sub InsertData2()
    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    'Dim rsSourceQD As DAO.QueryDef
    Dim rsTarget As DAO.Recordset
    Dim strSQL As String
    Dim startDate As Date
    Dim endDate As Date
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double
    Dim strTableA As String
    
    ' 取得打開的資料庫
    Set db = CurrentDb
    
    ' 設定起始日期和結束日期
    startDate = #7/1/2018#   ' 起始日期
    endDate = #7/31/2018# ' 結束日期
    strTableA = "dbo_vwFullAll"
    
    ' 取得來源資料表的記錄集
    Set rsSource = db.OpenRecordset(strTableA)
    
    ' 取得目標資料表的記錄集
    Set rsTarget = db.OpenRecordset("0508fullall2011")
    
    ' 開始計時
    startTime = Timer
    
    ' 設定日期迴圈
    Do While startDate <= endDate
        ' 篩選來源資料表的資料
        strSQL = "SELECT * FROM " & strTableA & " WHERE AddDate >= #" & startDate & "# AND AddDate < #" & DateAdd("m", 1, startDate) & "#;"
        'rsSource.OpenRecordset strSQL
        'Set rsSourceQD = db.CreateQueryDef("TempAddQuery")
        
        'rsSourceQD.SQL = strSQL
        'rsSource.OpenRecordset (strSQL)
        '= rsSourceQD.OpenRecordset()
        'Set rsSource = rsSourceQD.OpenRecordset()
        Set rsSource = db.OpenRecordset(strSQL, dbOpenDynaset)

        ' 取得當前月份的筆數
        Dim recordCount As Long
        recordCount = rsSource.recordCount
        
        ' 將當前月份和年份的筆數寫入偵錯 (Debug) 區
        Debug.Print "年份: " & Format(startDate, "yyyy") & "，月份: " & Format(startDate, "mm") & "，筆數: " & recordCount
        Dim iCount As Integer
        iCount = 1
        Do While rsSource.EOF <> True
            ' 插入到目標資料表
            rsTarget.AddNew
            For i = 0 To rsTarget.Fields.Count - 1
                rsTarget.Fields(i).Value = rsSource.Fields(i).Value
            Next i
            DoEvents
            'rsTarget.Fields("DateColumn").Value = rsSource.Fields("DateColumn").Value
            'rsTarget.Fields("ValueColumn").Value = rsSource.Fields("ValueColumn").Value
            rsTarget.Update
            iCount = iCount + 1
            Debug.Print "wrote:" & rsSource.Fields(0).Value & rsSource.Fields(1).Value & iCount
            If rsSource.EOF <> True Then
                rsSource.MoveNext
            End If
        Loop
        
        ' 關閉記錄集
        'db.QueryDefs.Delete "TempAddQuery"
        'rsSourceQD.Close
        
        ' 移動到下個月份
        startDate = DateAdd("m", 1, startDate)
        DoEvents
    Loop
    
        rsSource.Close
        rsTarget.Close

    
    ' 結束計時
    endTime = Timer
    elapsedTime = endTime - startTime
    
    ' 釋放記憶體
    Set rsSource = Nothing
    Set rsTarget = Nothing
    Set db = Nothing
    
    
    MsgBox "資料已成功插入到資料表 B。" & vbCrLf & "執行時間：" & elapsedTime & " 秒"
End Sub

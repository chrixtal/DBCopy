Attribute VB_Name = "BatchTransfer"
Option Compare Database

Sub InsertMultipleRecords()
    Dim db As DAO.Database
    'dim db2 as ad
    Dim strSQL As String
    Dim startDate As Date
    Dim endDate As Date
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double
    Dim totalRecordsInserted As Long
    Dim TableSrc As String  'the data output here
    Dim TableDes As String  'the data input here
    
    TableSrc = "dbo_vwFullAll"
    TableDes = "history"
    
    ' 取得打開的資料庫
    Set db = CurrentDb
    
    ' 設定起始日期和結束日期
'    startDate = #3/1/2009#   ' 起始日期
    startDate = #1/1/2018#   ' 起始日期
    endDate = #12/31/2018# ' 結束日期
    
    ' 開始計時
    startTime = Timer
    
    ' 設定日期迴圈
    Do While startDate <= endDate
        ' 構建要插入的資料的 SQL 查詢
        strSQL = "INSERT INTO " & TableDes & " (idx,cusna,tel,cellphone1,memo2,addr,zipcode,bday,sex,custype,branchname,receiptno,sel_no,reg_bm,emp_c,rec_m1,rec_m2,per_c,srem,pris_l,dist,arc,d_way,state,branchid,expr1,itemno,rdr_l,rdr_r,rsc_l,rsc_r,x_l,x_r,gl_no1,isnog1,gl_no2,isnog2,frame_no,isnogf,l_s_p,l_f_p,l_r_p,r_s_p,r_f_p,r_r_p,f_s_p,f_f_p,f_r_p,adddate,isget,getdate,gift,quantity,gl_no1name,gl_no2name,frame_name1,frame_name2) " & _
                 "SELECT idx,cusna,tel,cellphone1,memo2,addr,zipcode,bday,sex,custype,branchname,receiptno,sel_no,reg_bm,emp_c,rec_m1,rec_m2,per_c,srem,pris_l,dist,arc,d_way,state,branchid,expr1,itemno,rdr_l,rdr_r,rsc_l,rsc_r,x_l,x_r,gl_no1,isnog1,gl_no2,isnog2,frame_no,isnogf,l_s_p,l_f_p,l_r_p,r_s_p,r_f_p,r_r_p,f_s_p,f_f_p,f_r_p,adddate,isget,getdate,gift,quantity,gl_no1name,gl_no2name,frame_name1,frame_name2 " & _
                 "FROM " & TableSrc & _
                 " WHERE AddDate >= #" & startDate & "# AND AddDate < #" & DateAdd("m", 1, startDate) & "#;"
        
        ' 執行 SQL 查詢
        Dim recordsInserted As Long
        Dim ExeSTime As Double 'Execute Start Time and End Time
        Dim ExeETime As Double
        Dim ExePTime As Double 'Execute Period Time
        DoEvents
        ExeSTime = Timer
        db.Execute (strSQL)
        ExeETime = Timer
        ExePTime = ExeETime - ExeSTime
        recordsInserted = db.RecordsAffected
        Debug.Print startDate & " Total recound: " & db.RecordsAffected & " added, used time(min): " & ExePTime
        DoEvents
        totalRecordsInserted = totalRecordsInserted + recordsInserted
        
        ' 移動到下個月份
        startDate = DateAdd("m", 1, startDate)
    Loop
    
    ' 結束計時
    endTime = Timer
    elapsedTime = endTime - startTime
    
    ' 釋放記憶體
    Set db = Nothing
    
    ' 顯示統計資訊
    Dim strMsg As String
    strMsg = "Data insert                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   sucess to " & TableDes & vbCrLf & _
           "Executive Time: " & elapsedTime & " sec" & vbCrLf & _
           "Total insert: " & totalRecordsInserted & " records"
    MsgBox strMsg
    Debug.Print strMsg
           
End Sub


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
    
    ' ���o���}����Ʈw
    Set db = CurrentDb
    
    ' �]�w�_�l����M�������
    startDate = #7/1/2018#   ' �_�l���
    endDate = #7/31/2018# ' �������
    strTableA = "dbo_vwFullAll"
    
    ' ���o�ӷ���ƪ��O����
    Set rsSource = db.OpenRecordset(strTableA)
    
    ' ���o�ؼи�ƪ��O����
    Set rsTarget = db.OpenRecordset("0508fullall2011")
    
    ' �}�l�p��
    startTime = Timer
    
    ' �]�w����j��
    Do While startDate <= endDate
        ' �z��ӷ���ƪ����
        strSQL = "SELECT * FROM " & strTableA & " WHERE AddDate >= #" & startDate & "# AND AddDate < #" & DateAdd("m", 1, startDate) & "#;"
        'rsSource.OpenRecordset strSQL
        'Set rsSourceQD = db.CreateQueryDef("TempAddQuery")
        
        'rsSourceQD.SQL = strSQL
        'rsSource.OpenRecordset (strSQL)
        '= rsSourceQD.OpenRecordset()
        'Set rsSource = rsSourceQD.OpenRecordset()
        Set rsSource = db.OpenRecordset(strSQL, dbOpenDynaset)

        ' ���o��e���������
        Dim recordCount As Long
        recordCount = rsSource.recordCount
        
        ' �N��e����M�~�������Ƽg�J���� (Debug) ��
        Debug.Print "�~��: " & Format(startDate, "yyyy") & "�A���: " & Format(startDate, "mm") & "�A����: " & recordCount
        Dim iCount As Integer
        iCount = 1
        Do While rsSource.EOF <> True
            ' ���J��ؼи�ƪ�
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
        
        ' �����O����
        'db.QueryDefs.Delete "TempAddQuery"
        'rsSourceQD.Close
        
        ' ���ʨ�U�Ӥ��
        startDate = DateAdd("m", 1, startDate)
        DoEvents
    Loop
    
        rsSource.Close
        rsTarget.Close

    
    ' �����p��
    endTime = Timer
    elapsedTime = endTime - startTime
    
    ' ����O����
    Set rsSource = Nothing
    Set rsTarget = Nothing
    Set db = Nothing
    
    
    MsgBox "��Ƥw���\���J���ƪ� B�C" & vbCrLf & "����ɶ��G" & elapsedTime & " ��"
End Sub

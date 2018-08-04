Sub 导入数据2()
    Dim cnn As Object                '数据库连接
    Dim strcnn As String             'ACCESS连接语句
    Dim mydata As String            '数据库的完整路径和名称
    Dim mytable As String           '数据表名称
    Dim sql As String                  'sql查询语句
    Dim rs As Object                  '临时数据表纪录
    Dim i As Integer                  '循环数据变量（获取数据表字段）

    '1、连接数据库
    Set cnn = CreateObject("ADODB.Connection")
    mydata = ThisWorkbook.Path & "\测试数据库.accdb"

    Select Case Application.Version * 1    '设置连接字符串,根据版本创建连接
        Case Is <= 11
            strcnn = "Provider=Microsoft.Jet.Oledb.4.0;Jet OLEDB:Database Password='123456';Data Source=" & mydata
        Case Is >= 12
            strcnn = "Provider=Microsoft.ACE.OLEDB.12.0;Jet OLEDB:Database Password='123456';Data Source=" & mydata
    End Select
    
    cnn.Open strcnn    '打开数据库链接
 '2、设置sql查询语句
    mytable = "prediction"
    Set rs = CreateObject("ADODB.Recordset")

    sql = "select Seq,pkg,TotalPiece,TotalKgWeight,Description,ConsCompany,Conscity from prediction"
    'Seq,PKG#,TotalPiece,TotalKgWeight,Translation,Description,ConsCompany(EN),ConsCity
    Set rs = cnn.Execute(sql)    '执行查询，并将结果输出到记录集对象
    
    '3、复制数据库数据

    With ActiveSheet
        .Cells.ClearContents

        For i = 0 To rs.Fields.Count - 1    '填写标题
            .Cells(1, i + 1) = rs.Fields(i).Name
        Next i

        .Range("A2").CopyFromRecordset rs

        '.Cells.EntireColumn.AutoFit  '自动调整列宽
        '.Cells.EntireColumn.AutoFit  '自动调整列宽

    End With

    rs.Close
    cnn.Close
    Set rs = Nothing
    Set cnn = Nothing

End Sub

'VBA脚本功能：先将xxxx_b.xls中第二列数据清空，再将当前(或其他)目录的所
'有xxxx_a.xls合并到xxxx_b.xls中之后,按照户号升序排列并剔除信息完善的户
'和信息重复的情况，重复信息以xxxx_b.xls为准
'输入文件:xxxx_a.xls(承包方代表),xxxx_b.xls(身份证号码无效的情况) 
Sub formatTrans()
Dim myPath$, myFilea$, myFileb$, wa, wb As Workbook
Dim sa, sb As Worksheet	'a.sheet1,b.sheet1
Dim count, hstart, hend, enda, endb, filenum As Integer
Application.DisplayAlerts = False
Application.ScreenUpdating = False
myPath = ThisWorkbook.Path & "\"
myPath = InputBox("输入要处理的文件夹", "提示", myPath)
filenum = InputBox("输入末尾编号", "提示", 1)

For filecnt = 1 To filenum
	'添加前导0，有更好的方法？
    If filecnt > 0 And filecnt < 10 Then
        myFilea = "000" & filecnt & "_a.xls" 'filecnt.ToString("D4") & "_a.xls"
        myFileb = "000" & filecnt & "_b.xls" 'filecnt.ToString("D4") & "_b.xls"
    ElseIf filecnt >= 10 And filecnt < 100 Then
         myFilea = "00" & filecnt & "_a.xls" 'filecnt.ToString("D4") & "_a.xls"
        myFileb = "00" & filecnt & "_b.xls" 'filecnt.ToString("D4") & "_b.xls"
        
    ElseIf filecnt >= 100 And filecnt < 1000 Then
         myFilea = "0" & filecnt & "_a.xls" 'filecnt.ToString("D4") & "_a.xls"
        myFileb = "0" & filecnt & "_b.xls" 'filecnt.ToString("D4") & "_b.xls"
    ElseIf filecnt >= 1000 And filecnt < 10000 Then
         myFilea = "" & filecnt & "_a.xls" 'filecnt.ToString("D4") & "_a.xls"
        myFileb = "" & filecnt & "_b.xls" 'filecnt.ToString("D4") & "_b.xls"
    End If
	'调试输出，打开文件,获取行数
    Debug.Print filecnt & "、" & myFileb
    Set wa = Workbooks.Open(myPath & myFilea)
    Set sa = wa.Sheets(1)
    enda = [a65536].End(xlUp).Row
    
    Set wb = Workbooks.Open(myPath & myFileb)
    Set sb = wb.Sheets(1)
    endb = [a65536].End(xlUp).Row
	
    sb.Range("A2:N" & endb).Columns(2).Value = ""    '第二列置空
    sa.Range("A2:N" & enda).Copy sb.Range("A" & endb + 1)	'Copy a to b
    Workbooks(myFilea).Close True
    sb.UsedRange.Columns(13).Delete			'删除列
    sb.UsedRange.Columns(12).Delete
	'第二至N行升序排序
    With sb.Range("A2:L" & [a65536].End(xlUp).Row)   
            .Sort _
                Key1:=sb.Range("G1"), order1:=xlAscending 
    End With
	'删余
    Dim temp As String
    Dim endend As Integer
    endend = [a65536].End(xlUp).Row
    temp = sb.Cells(endend, 7).Value
    hstart = endend
    hend = endend
    For c = [a65536].End(xlUp).Row To 1 Step -1
        If sb.Cells(c, 7).Value <> temp Then
            hstart = c + 1
            If hend = hstart Then
                If sb.Cells(hstart, 2) <> "" Then
                    sb.Rows(hstart).Delete
                End If
            Else
                Dim x, y, cntt As Integer
                cntt = 0
                For f = hstart To hend
                    If sb.Cells(f, 4).Value = "户主" And cntt = 0 Then
                        x = f
                        cntt = cntt + 1
                    ElseIf sb.Cells(f, 4).Value = "户主" And cntt = 1 Then
                        y = f
                        cntt = cntt + 1
                    End If
                Next
                If cntt = 2 Then
                    If sb.Cells(x, 2).Value <> "" Then
                        sb.Rows(x).Delete
                    Else
                        sb.Rows(y).Delete
                    End If
                End If
            End If
            hend = c
            temp = sb.Cells(c, 7).Value
        End If
    Next
	'格式调整，以便打印
    With sb.Range("A1:L" & [a65536].End(xlUp).Row)   '全部[IV1].End(xlToLeft).Column
            .HorizontalAlignment = xlCenter    ' 所有单元格水平居中
            .Cells.VerticalAlignment = xlVAlignCenter   '所有单元格垂直居中
            .Borders.LineStyle = xlContinuous    '边框线
    End With
    sb.Columns(2).HorizontalAlignment = xlLeft    ' 左对齐
    With sb.Rows(1)    '第一行
            .RowHeight = 60
            .Font.Name = "Microsoft Sans Serif"
            .Font.Size = 10
            .WrapText = True
    End With
    With sb.Range("A2:L" & [a65536].End(xlUp).Row)   '第二至N行
            .Columns.AutoFit
            .Rows.RowHeight = 15
            '.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.Name = "Microsoft Sans Serif"
            .Font.Size = 10
            .WrapText = True
            .Sort _
                Key1:=Worksheets("Sheet1").Range("G1"), order1:=xlAscending '升序排序
            
    End With
    sb.Columns(2).ColumnWidth = 35     '单独设置列宽
    sb.Columns(10).ColumnWidth = 4     '单独设置列宽
    With sb.PageSetup
            .CenterHorizontally = True     '水平居中
            .CenterVertically = False      '取消垂直居中
            .PaperSize = xlPaperA4          'A4
            .LeftMargin = Application.CentimetersToPoints(1.5)      '左边距
            .RightMargin = Application.CentimetersToPoints(1.5)     '右边距
            .TopMargin = Application.CentimetersToPoints(1.1)       '上边距
            .BottomMargin = Application.CentimetersToPoints(1.1)    '下边距
            .FooterMargin = Application.CentimetersToPoints(0.5)    '页眉
            .HeaderMargin = Application.CentimetersToPoints(0.5)    '页脚
            .Orientation = xlLandscape '横向
            .Zoom = False				'所有列适合页宽,包括下面两行
            .FitToPagesWide = 1
            .FitToPagesTall = False
    End With
    'Sheets("sheet1").PrintOut
    Workbooks(myFileb).Close True
Next
Debug.Print "共完成" & filecnt - 1 & "个文件"
MsgBox ("共处理 " & filecnt - 1 & " 个文件")
Application.DisplayAlerts = True
Application.DisplayAlerts = True
End Sub


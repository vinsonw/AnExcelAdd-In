Attribute VB_Name = "阶段性报告制作v1"
'Date: 2016.8.1-2016.9.18
'Author: Vinson Wei
'Purpose: Automate the process of making a periodic report

Dim factoryinphase$, year$, startMonth$, endMonth$, currentUseWB As Workbook

Dim status As Integer

Dim HRC00, HRC01, HRC02, HRC03, HRC04, HRC05, HRC06, HRC07, HRC08, HRC09, HRC10, HRC11, HRC12, _
HRC13, HRC14, HRC14Dot5, HRC14Dot6, HRC15, HRC16, HRC17, HRC18, HRC19, HRC20, HRC21, HRC22, _
HRC23, HRC24, HRC25, HRC26, HRC27, HRC28, HRC29, HRC30, HRC31, HRC32, HRC33 As Integer


Sub mainInitializationInPhaseReport(ByRef status As Integer)
factoryinphase = InputBox(prompt:="请输入工厂代码：")

If factoryinphase = "" Then
status = 1
Exit Sub
End If

year = InputBox(prompt:="请输入本文件夹下月报所属年份，如2015")

If year = "" Then
status = 2
Exit Sub
End If

startMonth = InputBox(prompt:="请输入本年度、本文件夹下月报的起始月，如3（表示3月）")

If startMonth = "" Then
status = 3
Exit Sub
End If

endMonth = InputBox(prompt:="请输入本年度、本文件夹下月报的最后一个月，如12（表示12月）")

If endMonth = "" Then
status = 4
Exit Sub
End If

Set currentUseWB = ActiveWorkbook
On Error Resume Next
currentUseWB.Worksheets("Sheet2").Delete
currentUseWB.Worksheets("Sheet3").Delete
End Sub

Sub mainOpenCopyCloseOneiMonthlyReport(control As IRibbonControl)
'control As IRibbonControl
status = 0
Call mainInitializationInPhaseReport(status)

If status <> 0 Then
    Select Case status
        Case 1
            MsgBox "未输入工厂代码，点击确定退出月报数据导入"
        Case 2
            MsgBox "未输入年度，点击确定退出月报数据导入"
        Case 3
            MsgBox "未输入本年度月报起始月份，点击确定退出月报数据导入"
        Case 4
            MsgBox "未输入本年度月报最后一个月份，点击确定退出月报数据导入"
    End Select
Exit Sub
End If

Dim f$, WB As Workbook, pw$, loc$, sht As Worksheet, nameList

'写好汇总表表头
If currentUseWB.Worksheets("Sheet1").Range("H1").Value = "" Then

    currentUseWB.Worksheets("Sheet1").Range("A1").Value = "年月"
    currentUseWB.Worksheets("Sheet1").Range("B1").Value = "个案编号"
    currentUseWB.Worksheets("Sheet1").Range("C1").Value = "颜色"
    currentUseWB.Worksheets("Sheet1").Range("D1").Value = "事件类型"
    currentUseWB.Worksheets("Sheet1").Range("E1").Value = "问题分类"
    currentUseWB.Worksheets("Sheet1").Range("F1").Value = "沟通方式"
    currentUseWB.Worksheets("Sheet1").Range("G1").Value = "事主性别"
    currentUseWB.Worksheets("Sheet1").Range("H1").Value = "案件详述"
    currentUseWB.Worksheets("Sheet1").Range("I1").Value = "个案描述"
    rStart = 2
    lStart = 2
Else
    Set lastCell = currentUseWB.Worksheets("Sheet1").Range("H1").End(xlDown)
    rStart = lastCell.Row + 1
    lStart = rStart
End If

iRInSheet1 = rStart
iLInSheet1 = lStart
For iMonth = Val(startMonth) To Val(endMonth)

'查询月报文件，打开月报文件
f = Dir(currentUseWB.Path & "\*" & factoryinphase & "*" & iMonth & "月*" & "个案汇总.xls*")
loc = InStr(f, " ")
If loc = 0 Then loc = InStr(f, " ")
pw = Trim(Mid(f, 1, loc - 1))
Set WB = Workbooks.Open(Filename:=currentUseWB.Path & "\" & f, Password:=pw)


'将个案工作表表名加入到nameList，完成排序
Set nameList = CreateObject("System.Collections.ArrayList")
For Each sht In WB.Worksheets
    If sht.Name Like ("*" & factoryinphase & "*") Then
        nameList.Add sht.Name
    End If
Next

'先从月报的每张个案工作表中将案件详述搬运过来

For i = 0 To nameList.Count - 1
    WB.Worksheets(nameList(i)).Range("B5").Copy currentUseWB.Worksheets("Sheet1").Range("H" & iRInSheet1)
    With currentUseWB.Worksheets("Sheet1").Range("H" & iRInSheet1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    iRInSheet1 = iRInSheet1 + 1
Next

'再从Quesiton Sheet表中将个案编号、颜色、事件类型、问题分类一次都搬运过来
WB.Worksheets("Question Sheet").Range("A2", "D" & (nameList.Count + 1)).Copy _
currentUseWB.Worksheets("Sheet1").Range("B" & iLInSheet1)


'从Case Sheet表中将沟通方式、事主性别搬运过来
WB.Worksheets("Case Sheet").Range("F2", "G" & (nameList.Count + 1)).Copy _
currentUseWB.Worksheets("Sheet1").Range("F" & iLInSheet1)

'写入年月
For i = iLInSheet1 To (iRInSheet1 - 1)
currentUseWB.Worksheets("Sheet1").Range("A" & i).Value = "'" & year & "年" & iMonth & "月"
iLInSheet1 = iLInSheet1 + 1
Next

endCode = currentUseWB.Worksheets("Sheet1").Range("B" & (iLInSheet1 - 1)).Value

If Trim(endCode) <> nameList(nameList.Count - 1) Then
    MsgBox "注意，现在发生了一个罕见的情况，我猜" & year & "年" & iMonth & "月的月报中" _
    & "的某个个案工作表里面有两个个案（但共用一个个案编号），这" & "导致了导入时发生错位，点击确定，程序将尝试自动解决错位。"
    GoTo iLNotEqualToiR
End If

'正常的在这里开始
iLEqualToiR:
iLInSheet1 = iRInSheet1

WB.Close False

Next

With currentUseWB.Worksheets("Sheet1").Range("A1", "G" & (iLInSheet1 - 1))
        .RowHeight = 25
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With

currentUseWB.Worksheets("Sheet1").Columns("A:I").AutoFit
currentUseWB.Worksheets("Sheet1").Columns("H").ColumnWidth = 45
'MsgBox "iLInSheet1:" & iLInSheet1 & "iRInSheet1:" & iRInSheet1
MsgBox year & "年度月报导入完毕，请点击确定，如果还有下一年度的月报需要导入，请先复制本Excel文件至这部分月报所在文件夹，重复导入步骤"

Exit Sub
iLNotEqualToiR:
currentUseWB.Worksheets("Sheet1").Activate
stRow = iRInSheet1 - nameList.Count
Set range00 = currentUseWB.Worksheets("Sheet1").Cells(stRow, 1)
range00.Resize(nameList.Count + 2, 8).Clear


iRInSheet1 = iRInSheet1 - nameList.Count

startRowInThis = iRInSheet1

storeSheet = ""
For i = 0 To nameList.Count - 1
       WB.Worksheets(nameList(i)).Range("B5").Copy currentUseWB.Worksheets("Sheet1").Range("H" & iRInSheet1)
   If Len(Trim(WB.Worksheets(nameList(i)).Range("A6").Value)) <> 0 Then
        storeSheet = nameList(i)
       WB.Worksheets(nameList(i)).Range("B6").Copy currentUseWB.Worksheets("Sheet1").Range("H" & (iRInSheet1 + 1))
   End If
    With currentUseWB.Worksheets("Sheet1").Range("H" & iRInSheet1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    If Len(Trim(WB.Worksheets(nameList(i)).Range("A6").Value)) <> 0 Then
        iRInSheet1 = iRInSheet1 + 2
    Else
         iRInSheet1 = iRInSheet1 + 1
    End If
Next

'再从Quesiton Sheet表中将个案编号、颜色、事件类型、问题分类一次都搬运过来
endRowInThis = WB.Worksheets("Question Sheet").Range("A2").End(xlDown).Row
WB.Worksheets("Question Sheet").Range("A2").Resize(endRowInThis - 1, 4).Copy _
currentUseWB.Worksheets("Sheet1").Range("B" & startRowInThis)

'======================2016年9月7日00:11:28工作到这里=====================
'谈到storeSheet，即存下多事件的工作表名，然后逐个不用循环填入沟通方式、事主性别

'从Case Sheet表中将沟通方式、事主性别搬运过来
startRowInThis0 = startRowInThis
For i = 2 To nameList.Count + 1
    pTroubleCase = factoryinphase & "-" & WB.Worksheets("Case Sheet").Cells(i, 4).Value
    If storeSheet = pTroubleCase Then
     WB.Worksheets("Case Sheet").Cells(i, 6).Resize(1, 2).Copy currentUseWB.Worksheets("Sheet1").Cells(startRowInThis0, 6).Resize(1, 2)
     WB.Worksheets("Case Sheet").Cells(i, 6).Resize(1, 2).Copy currentUseWB.Worksheets("Sheet1").Cells(startRowInThis0 + 1, 6).Resize(1, 2)
     startRowInThis0 = startRowInThis0 + 2
     Else
        WB.Worksheets("Case Sheet").Cells(i, 6).Resize(1, 2).Copy currentUseWB.Worksheets("Sheet1").Cells(startRowInThis0, 6).Resize(1, 2)
        startRowInThis0 = startRowInThis0 + 1
    End If
Next


'写入年月，已修改
For i = startRowInThis To (iRInSheet1 - 1)
currentUseWB.Worksheets("Sheet1").Range("A" & i).Value = "'" & year & "年" & iMonth & "月"
iLInSheet1 = iLInSheet1 + 1
Next
MsgBox "Ok，错位问题已解决，点击继续程序运行。"
GoTo iLEqualToiR


End Sub
Sub mainCreatePivotTable(control As IRibbonControl)
On Error Resume Next
    Application.DisplayAlerts = False
    currentUseWB.Sheets("PivotSheet").Delete
    Application.DisplayAlerts = False
On Error GoTo 0


Dim pivotSheet As Worksheet, rg As Range, endCell As Range
Set currentUseWB = ActiveWorkbook
'创建数据源
Set pivotSheet = currentUseWB.Worksheets.Add(after:=currentUseWB.Worksheets("Sheet1"))
pivotSheet.Name = "PivotSheet"
Set endCell = currentUseWB.Worksheets("Sheet1").Range("E2").End(xlDown)
currentUseWB.Worksheets("Sheet1").Activate
endCell.Select
Set rg = Range(Range("D1"), endCell)

 currentUseWB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rg, Version:=xlPivotTableVersion10).CreatePivotTable _
        TableDestination:="PivotSheet!R3C1", TableName:="PivotTable", DefaultVersion:= _
        xlPivotTableVersion10
    Worksheets("PivotSheet").Select
    currentUseWB.ShowPivotTableFieldList = True
    With currentUseWB.Worksheets("PivotSheet").PivotTables("PivotTable").PivotFields("事件类型")
        .Orientation = xlRowField
        .Position = 1
    End With
        currentUseWB.Worksheets("PivotSheet").PivotTables("PivotTable").AddDataField currentUseWB.Worksheets("PivotSheet").PivotTables("PivotTable" _
        ).PivotFields("问题分类"), "计数项：问题分类", xlCount
    With currentUseWB.Worksheets("PivotSheet").PivotTables("PivotTable").PivotFields("问题分类")
        .Orientation = xlColumnField
        .Position = 1
    End With

    currentUseWB.ShowPivotTableFieldList = False
    
    MsgBox "数据透视表生成完毕"
End Sub

Sub mainCreatePieChart(control As IRibbonControl)
Set currentUseWB = ActiveWorkbook
' 创建数据源


On Error Resume Next
    Application.DisplayAlerts = False
    currentUseWB.Sheets("PieChart").Delete
    currentUseWB.Sheets("Gender").Delete
    currentUseWB.Sheets("InMethod").Delete
    currentUseWB.Sheets("SourceData").Delete
    Application.DisplayAlerts = False
On Error GoTo 0
' 饼图数据
Dim srcSheeet As Worksheet
Dim d, d1, d2 As Object, arr, i&
Set d = CreateObject("scripting.dictionary")
currentUseWB.Worksheets("Sheet1").Activate
arr = currentUseWB.Worksheets("Sheet1").Range([E2], [E65536].End(xlUp))
For i = 1 To UBound(arr)
    If arr(i, 1) <> "" Then d(arr(i, 1)) = ""
Next

Set srcSheet = currentUseWB.Worksheets.Add(after:=currentUseWB.Worksheets(currentUseWB.Worksheets.Count))
srcSheet.Name = "SourceData"

currentUseWB.Worksheets("SourceData").[A2].Resize(d.Count) = Application.Transpose(d.keys)

currentUseWB.Worksheets("SourceData").[B1].Value = "个案问题"

currentUseWB.Worksheets("Sheet1").Activate
Set arr = currentUseWB.Worksheets("Sheet1").Range([E2], [E65536].End(xlUp))

For i = 2 To d.Count + 1
currentUseWB.Worksheets("SourceData").Range("B" & i) = _
WorksheetFunction.CountIf(arr, currentUseWB.Worksheets("SourceData").Range("A" & i))
Next

currentUseWB.Worksheets("SourceData").Sort.SortFields.Clear
   currentUseWB.Worksheets("SourceData").Sort.SortFields.Add Key:=Range("B2") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With currentUseWB.Worksheets("SourceData").Sort
        .SetRange Range("A2:B" & (d.Count + 1))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


' 性别比例柱状图
genderStartRow = d.Count + 3

Set d1 = CreateObject("scripting.dictionary")
arr = currentUseWB.Worksheets("Sheet1").Range([G2], [G65536].End(xlUp))
For i = 1 To UBound(arr)
    If arr(i, 1) <> "" Then d1(arr(i, 1)) = ""
Next

currentUseWB.Worksheets("SourceData").Range("B" & genderStartRow) = "性别"
currentUseWB.Worksheets("SourceData").Range("A" & (genderStartRow + 1)).Resize(d1.Count) = Application.Transpose(d1.keys)

currentUseWB.Worksheets("Sheet1").Activate
Set arr = currentUseWB.Worksheets("Sheet1").Range([G2], [G65536].End(xlUp))

For i = genderStartRow + 1 To genderStartRow + d1.Count
currentUseWB.Worksheets("SourceData").Range("B" & i) = _
WorksheetFunction.CountIf(arr, currentUseWB.Worksheets("SourceData").Range("A" & i))
Next

'接入方式数据
inMethodStartRow = genderStartRow + d1.Count + 2

Set d2 = CreateObject("scripting.dictionary")
arr = currentUseWB.Worksheets("Sheet1").Range([F2], [F65536].End(xlUp))
For i = 1 To UBound(arr)
    If arr(i, 1) <> "" Then d2(arr(i, 1)) = ""
Next

currentUseWB.Worksheets("SourceData").Range("B" & inMethodStartRow) = "接入方式"
currentUseWB.Worksheets("SourceData").Range("A" & (inMethodStartRow + 1)).Resize(d2.Count) = Application.Transpose(d2.keys)

currentUseWB.Worksheets("Sheet1").Activate
Set arr = currentUseWB.Worksheets("Sheet1").Range([F2], [F65536].End(xlUp))

For i = inMethodStartRow + 1 To inMethodStartRow + d2.Count
currentUseWB.Worksheets("SourceData").Range("B" & i) = _
WorksheetFunction.CountIf(arr, currentUseWB.Worksheets("SourceData").Range("A" & i))
Next



'做出符合要求的饼图
If d.Count > 4 Then
    If currentUseWB.Worksheets("SourceData").[B5] = currentUseWB.Worksheets("SourceData").[B6] Then
        If currentUseWB.Worksheets("SourceData").[B6] = currentUseWB.Worksheets("SourceData").[B7] Then
            If currentUseWB.Worksheets("SourceData").[B7] = currentUseWB.Worksheets("SourceData").[B8] Then
                If currentUseWB.Worksheets("SourceData").[B8] = currentUseWB.Worksheets("SourceData").[B9] Then
                    endRow = 9
                Else
                    endRow = 8
                End If
            Else
                endRow = 7
            End If
        Else
            endRow = 6
        End If
    Else
        endRow = 5
    End If
Else
    endRow = d.Count + 1
End If

currentUseWB.Worksheets("SourceData").Range("s2").Value = endRow
currentUseWB.Worksheets("SourceData").Range("s1").Value = "饼图中最后一个问题分类所在行数(endRow)"

sumOfPie = WorksheetFunction.Sum(currentUseWB.Worksheets("SourceData").Range("B2", "B" & endRow))

For i = 2 To endRow
    currentUseWB.Worksheets("SourceData").Range("B" & i) = _
    currentUseWB.Worksheets("SourceData").Range("B" & i) / sumOfPie
Next
currentUseWB.Worksheets("SourceData").Range("B2", "B" & endRow).NumberFormatLocal = "0.00%"


'===========================================
'此处统计Word文档第四部分所需数据
'数据位置在SourceData工作表中

'写好表格框架
realEndRow = currentUseWB.Worksheets("SourceData").[A2].End(xlDown).Row
For i = 2 To realEndRow

    For j = 3 To 17 Step 2
        If j = 3 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "咨询"
        If j = 5 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "占咨询百分比"
        If j = 7 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "心理"
        If j = 9 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "占心理百分比"
        If j = 11 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "投诉"
        If j = 13 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "占投诉百分比"
        If j = 15 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = currentUseWB.Worksheets("SourceData").Cells(i, 1).Value
        If j = 17 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "占总个案百分比"
    Next

Next



'统计到endRow为止各问题分类所在事件类型

For i = 2 To realEndRow

    For j = 4 To 16 Step 4
        currentUseWB.Worksheets("SourceData").Cells(i, j).Value = _
        Application.WorksheetFunction.CountIfs(currentUseWB.Worksheets("Sheet1").Range("D:D"), currentUseWB.Worksheets("SourceData").Cells(i, j - 1).Value, _
        currentUseWB.Worksheets("Sheet1").Range("E:E"), currentUseWB.Worksheets("SourceData").Cells(i, 1).Value)
    Next
    

Next

'=========================
ctGreen = 0
ctYellow = 0
ctRed = 0
ctZixun = 0
ctXinli = 0
ctTousu = 0

With currentUseWB.Worksheets("Sheet1")

For i = 2 To .[A65536].End(xlUp).Row
    If Trim(.Range("C" & i).Value) = "绿" Then ctGreen = ctGreen + 1
    If Trim(.Range("C" & i).Value) = "黄" Then ctYellow = ctYellow + 1
    If Trim(.Range("C" & i).Value) = "红" Then ctRed = ctRed + 1
    If Trim(.Range("D" & i).Value) = "咨询" Then ctZixun = ctZixun + 1
    If Trim(.Range("D" & i).Value) = "投诉" Then ctTousu = ctTousu + 1
    If Trim(.Range("D" & i).Value) = "心理" Then ctXinli = ctXinli + 1
Next

End With

'=========================


For i = 2 To realEndRow

    For j = 6 To 14 Step 4
    
    If j = 6 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, j - 2).Value / ctZixun, "0.00%")
    If j = 10 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, j - 2).Value / ctXinli, "0.00%")
    If j = 14 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, j - 2).Value / ctTousu, "0.00%")
    Next
Next

Set hiDict = CreateObject("scripting.dictionary")
hiDict.Add "咨询", ctZixun
hiDict.Add "心理", ctXinli
hiDict.Add "投诉", ctTousu

numOfAllCases = currentUseWB.Worksheets("Sheet1").[A65536].End(xlUp).Row - 1
For i = 2 To realEndRow

        currentUseWB.Worksheets("SourceData").Cells(i, 16).Value = _
        Application.WorksheetFunction.CountIfs(currentUseWB.Worksheets("Sheet1").Range("E:E"), currentUseWB.Worksheets("SourceData").Cells(i, 15).Value)
    currentUseWB.Worksheets("SourceData").Cells(i, 18).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, 16).Value / numOfAllCases, "0.00%")
    
Next



'===========================================

Set pieChart = currentUseWB.Charts.Add(after:=currentUseWB.Worksheets("SourceData"))
pieChart.Name = "PieChart"
pieChart.ChartType = xl3DPieExploded
pieChart.SetSourceData Source:=currentUseWB.Sheets("SourceData").Range("A1:B" & endRow)
pieChart.SeriesCollection(1).Select
pieChart.SeriesCollection(1).ApplyDataLabels
pieChart.PlotArea.Select
pieChart.ChartArea.Select

Set barChart = currentUseWB.Charts.Add(after:=currentUseWB.Worksheets("SourceData"))
barChart.Name = "Gender"
barChart.ChartType = xlColumnClustered
barChart.SetSourceData Source:=currentUseWB.Sheets("SourceData").Range("A" & genderStartRow & ":B" & (genderStartRow + d1.Count))
barChart.SeriesCollection(1).ApplyDataLabels

Set barChart1 = currentUseWB.Charts.Add(after:=currentUseWB.Worksheets("SourceData"))
barChart1.Name = "InMethod"
barChart1.ChartType = xlColumnClustered
barChart1.SetSourceData Source:=currentUseWB.Sheets("SourceData").Range("A" & inMethodStartRow & ":B" & (inMethodStartRow + d2.Count))
barChart1.SeriesCollection(1).ApplyDataLabels

currentUseWB.Worksheets("Sheet1").Activate
currentUseWB.Worksheets("Sheet1").[A1].Select

MsgBox "成功生成了一幅饼图和两幅柱状图，请注意如果两幅柱状图中数据只有一列（比如只有男或者只有QQ接入），则柱状图的格式可能有问题，需要手工修改格式"

End Sub
Sub mainCreatePieChart1()
Set currentUseWB = ActiveWorkbook
' 创建数据源


On Error Resume Next
    Application.DisplayAlerts = False
    currentUseWB.Sheets("PieChart").Delete
    currentUseWB.Sheets("Gender").Delete
    currentUseWB.Sheets("InMethod").Delete
    currentUseWB.Sheets("SourceData").Delete
    Application.DisplayAlerts = False
On Error GoTo 0
' 饼图数据
Dim srcSheeet As Worksheet
Dim d, d1, d2 As Object, arr, i&
Set d = CreateObject("scripting.dictionary")
currentUseWB.Worksheets("Sheet1").Activate
arr = currentUseWB.Worksheets("Sheet1").Range([E2], [E65536].End(xlUp))
For i = 1 To UBound(arr)
    If arr(i, 1) <> "" Then d(arr(i, 1)) = ""
Next

Set srcSheet = currentUseWB.Worksheets.Add(after:=currentUseWB.Worksheets(currentUseWB.Worksheets.Count))
srcSheet.Name = "SourceData"

currentUseWB.Worksheets("SourceData").[A2].Resize(d.Count) = Application.Transpose(d.keys)

currentUseWB.Worksheets("SourceData").[B1].Value = "个案问题"

currentUseWB.Worksheets("Sheet1").Activate
Set arr = currentUseWB.Worksheets("Sheet1").Range([E2], [E65536].End(xlUp))

For i = 2 To d.Count + 1
currentUseWB.Worksheets("SourceData").Range("B" & i) = _
WorksheetFunction.CountIf(arr, currentUseWB.Worksheets("SourceData").Range("A" & i))
Next

currentUseWB.Worksheets("SourceData").Sort.SortFields.Clear
   currentUseWB.Worksheets("SourceData").Sort.SortFields.Add Key:=Range("B2") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With currentUseWB.Worksheets("SourceData").Sort
        .SetRange Range("A2:B" & (d.Count + 1))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


' 性别比例柱状图
genderStartRow = d.Count + 3

Set d1 = CreateObject("scripting.dictionary")
arr = currentUseWB.Worksheets("Sheet1").Range([G2], [G65536].End(xlUp))
For i = 1 To UBound(arr)
    If arr(i, 1) <> "" Then d1(arr(i, 1)) = ""
Next

currentUseWB.Worksheets("SourceData").Range("B" & genderStartRow) = "性别"
currentUseWB.Worksheets("SourceData").Range("A" & (genderStartRow + 1)).Resize(d1.Count) = Application.Transpose(d1.keys)

currentUseWB.Worksheets("Sheet1").Activate
Set arr = currentUseWB.Worksheets("Sheet1").Range([G2], [G65536].End(xlUp))

For i = genderStartRow + 1 To genderStartRow + d1.Count
currentUseWB.Worksheets("SourceData").Range("B" & i) = _
WorksheetFunction.CountIf(arr, currentUseWB.Worksheets("SourceData").Range("A" & i))
Next

'接入方式数据
inMethodStartRow = genderStartRow + d1.Count + 2

Set d2 = CreateObject("scripting.dictionary")
arr = currentUseWB.Worksheets("Sheet1").Range([F2], [F65536].End(xlUp))
For i = 1 To UBound(arr)
    If arr(i, 1) <> "" Then d2(arr(i, 1)) = ""
Next

currentUseWB.Worksheets("SourceData").Range("B" & inMethodStartRow) = "接入方式"
currentUseWB.Worksheets("SourceData").Range("A" & (inMethodStartRow + 1)).Resize(d2.Count) = Application.Transpose(d2.keys)

currentUseWB.Worksheets("Sheet1").Activate
Set arr = currentUseWB.Worksheets("Sheet1").Range([F2], [F65536].End(xlUp))

For i = inMethodStartRow + 1 To inMethodStartRow + d2.Count
currentUseWB.Worksheets("SourceData").Range("B" & i) = _
WorksheetFunction.CountIf(arr, currentUseWB.Worksheets("SourceData").Range("A" & i))
Next



'做出符合要求的饼图
If d.Count > 4 Then
    If currentUseWB.Worksheets("SourceData").[B5] = currentUseWB.Worksheets("SourceData").[B6] Then
        If currentUseWB.Worksheets("SourceData").[B6] = currentUseWB.Worksheets("SourceData").[B7] Then
            If currentUseWB.Worksheets("SourceData").[B7] = currentUseWB.Worksheets("SourceData").[B8] Then
                If currentUseWB.Worksheets("SourceData").[B8] = currentUseWB.Worksheets("SourceData").[B9] Then
                    endRow = 9
                Else
                    endRow = 8
                End If
            Else
                endRow = 7
            End If
        Else
            endRow = 6
        End If
    Else
        endRow = 5
    End If
Else
    endRow = d.Count + 1
End If

currentUseWB.Worksheets("SourceData").Range("s2").Value = endRow
currentUseWB.Worksheets("SourceData").Range("s1").Value = "饼图中最后一个问题分类所在行数(endRow)"

sumOfPie = WorksheetFunction.Sum(currentUseWB.Worksheets("SourceData").Range("B2", "B" & endRow))

For i = 2 To endRow
    currentUseWB.Worksheets("SourceData").Range("B" & i) = _
    currentUseWB.Worksheets("SourceData").Range("B" & i) / sumOfPie
Next
currentUseWB.Worksheets("SourceData").Range("B2", "B" & endRow).NumberFormatLocal = "0.00%"


'===========================================
'此处统计Word文档第四部分所需数据
'数据位置在SourceData工作表中

'写好表格框架
realEndRow = currentUseWB.Worksheets("SourceData").[A2].End(xlDown).Row
For i = 2 To realEndRow

    For j = 3 To 17 Step 2
        If j = 3 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "咨询"
        If j = 5 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "占咨询百分比"
        If j = 7 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "心理"
        If j = 9 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "占心理百分比"
        If j = 11 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "投诉"
        If j = 13 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "占投诉百分比"
        If j = 15 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = currentUseWB.Worksheets("SourceData").Cells(i, 1).Value
        If j = 17 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = "占总个案百分比"
    Next

Next



'统计到endRow为止各问题分类所在事件类型

For i = 2 To realEndRow

    For j = 4 To 16 Step 4
        currentUseWB.Worksheets("SourceData").Cells(i, j).Value = _
        Application.WorksheetFunction.CountIfs(currentUseWB.Worksheets("Sheet1").Range("D:D"), currentUseWB.Worksheets("SourceData").Cells(i, j - 1).Value, _
        currentUseWB.Worksheets("Sheet1").Range("E:E"), currentUseWB.Worksheets("SourceData").Cells(i, 1).Value)
    Next
    

Next

'=========================
ctGreen = 0
ctYellow = 0
ctRed = 0
ctZixun = 0
ctXinli = 0
ctTousu = 0

With currentUseWB.Worksheets("Sheet1")

For i = 2 To .[A65536].End(xlUp).Row
    If Trim(.Range("C" & i).Value) = "绿" Then ctGreen = ctGreen + 1
    If Trim(.Range("C" & i).Value) = "黄" Then ctYellow = ctYellow + 1
    If Trim(.Range("C" & i).Value) = "红" Then ctRed = ctRed + 1
    If Trim(.Range("D" & i).Value) = "咨询" Then ctZixun = ctZixun + 1
    If Trim(.Range("D" & i).Value) = "投诉" Then ctTousu = ctTousu + 1
    If Trim(.Range("D" & i).Value) = "心理" Then ctXinli = ctXinli + 1
Next

End With

'=========================


For i = 2 To realEndRow

    For j = 6 To 14 Step 4
    
    If j = 6 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, j - 2).Value / ctZixun, "0.00%")
    If j = 10 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, j - 2).Value / ctXinli, "0.00%")
    If j = 14 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, j - 2).Value / ctTousu, "0.00%")
    Next
Next

Set hiDict = CreateObject("scripting.dictionary")
hiDict.Add "咨询", ctZixun
hiDict.Add "心理", ctXinli
hiDict.Add "投诉", ctTousu

numOfAllCases = currentUseWB.Worksheets("Sheet1").[A65536].End(xlUp).Row - 1
For i = 2 To realEndRow

        currentUseWB.Worksheets("SourceData").Cells(i, 16).Value = _
        Application.WorksheetFunction.CountIfs(currentUseWB.Worksheets("Sheet1").Range("E:E"), currentUseWB.Worksheets("SourceData").Cells(i, 15).Value)
    currentUseWB.Worksheets("SourceData").Cells(i, 18).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, 16).Value / numOfAllCases, "0.00%")
    
Next



'===========================================

Set pieChart = currentUseWB.Charts.Add(after:=currentUseWB.Worksheets("SourceData"))
pieChart.Name = "PieChart"
pieChart.ChartType = xl3DPieExploded
pieChart.SetSourceData Source:=currentUseWB.Sheets("SourceData").Range("A1:B" & endRow)
pieChart.SeriesCollection(1).Select
pieChart.SeriesCollection(1).ApplyDataLabels
pieChart.PlotArea.Select
pieChart.ChartArea.Select

Set barChart = currentUseWB.Charts.Add(after:=currentUseWB.Worksheets("SourceData"))
barChart.Name = "Gender"
barChart.ChartType = xlColumnClustered
barChart.SetSourceData Source:=currentUseWB.Sheets("SourceData").Range("A" & genderStartRow & ":B" & (genderStartRow + d1.Count))
barChart.SeriesCollection(1).ApplyDataLabels

Set barChart1 = currentUseWB.Charts.Add(after:=currentUseWB.Worksheets("SourceData"))
barChart1.Name = "InMethod"
barChart1.ChartType = xlColumnClustered
barChart1.SetSourceData Source:=currentUseWB.Sheets("SourceData").Range("A" & inMethodStartRow & ":B" & (inMethodStartRow + d2.Count))
barChart1.SeriesCollection(1).ApplyDataLabels


MsgBox "成功生成了一幅饼图和两幅柱状图，请注意如果两幅柱状图中数据只有一列（比如只有男或者只有QQ接入），则柱状图的格式可能有问题，需要手工修改格式"

End Sub

Sub mainCreateWordDoc(control As IRibbonControl)


mainCreatePieChart1

Set currentUseWB = ActiveWorkbook

Dim f$
f = Dir(currentUseWB.Path & "\热线阶段性报告模版.docx")
If f = "" Then f = Dir(currentUseWB.Path & "\热线阶段性报告模版.doc")
If f = "" Then
    MsgBox "程序在本Excel文件所在文件夹下没有找到指定的阶段性报告模板文件【热线阶段性报告模版.docx】，点击确定退出程序"
    Exit Sub
End If


Dim sht As Worksheet

Application.DisplayAlerts = False
On Error Resume Next
currentUseWB.Worksheets("ToWordDoc").Delete
currentUseWB.Worksheets("pie").Delete
currentUseWB.Worksheets("dgender").Delete
currentUseWB.Worksheets("dinmethod").Delete
On Error GoTo 0

Application.DisplayAlerts = True

Set sht = currentUseWB.Worksheets.Add(after:=currentUseWB.Sheets(currentUseWB.Sheets.Count))

sht.Name = "ToWordDoc"

With currentUseWB.Worksheets("ToWordDoc")
    .[A1].Value = "位置说明"
    .[A1].Interior.Color = vbYellow
    .[A2].Value = "HRC编号"
    .[A2].Interior.Color = vbGreen
    .[A3].Value = "值"
    .[A3].Interior.Color = vbCyan
    .[A4].Value = "位置说明"
    .[A4].Interior.Color = vbYellow
    .[A5].Value = "HRC编号"
    .[A5].Interior.Color = vbGreen
    .[A6].Value = "值"
    .[A6].Interior.Color = vbCyan
    .[A7].Value = "位置说明"
    .[A7].Interior.Color = vbYellow
    .[A8].Value = "HRC编号"
    .[A8].Interior.Color = vbGreen
    .[A9].Value = "值"
    .[A9].Interior.Color = vbCyan
    
    .[B1].Value = "工厂代号"
    .[B2].Value = "HRC00"
    .[C1].Value = "起始年"
    .[C2].Value = "HRC01"
    .[d1].Value = "起始月"
    .[d2].Value = "HRC02"
    .[E1].Value = "终止年"
    .[E2].Value = "HRC03"
    .[F1].Value = "终止月"
    .[F2].Value = "HRC04"
    .[G1].Value = "起始年"
    .[G2].Value = "HRC05"
    .[H1].Value = "起始月"
    .[H2].Value = "HRC06"
    .[I1].Value = "终止年"
    .[I2].Value = "HRC07"
    .[J1].Value = "终止月"
    .[J2].Value = "HRC08"
    .[K1].Value = "个案总数"
    .[K2].Value = "HRC09"
    .[L1].Value = "问题分类总数"
    .[L2].Value = "HRC10"
    .[M1].Value = "咨询类个案数"
    .[M2].Value = "HRC11"
    .[N1].Value = "占个案总数百分比"
    .[N2].Value = "HRC12"
    .[O1].Value = "心理类个案数"
    .[O2].Value = "HRC13"
    .[P1].Value = "占个案总数百分比"
    .[P2].Value = "HRC14"
    .[Q1].Value = "投诉类个案数"
    .[Q2].Value = "HRC14Dot5"
    .[R1].Value = "占个案总数百分比"
    .[R2].Value = "HRC14Dot6"
    .[B4].Value = "绿色个案数"
    .[B5].Value = "HRC15"
    .[C4].Value = "黄色个案数"
    .[C5].Value = "HRC16"
    .[D4].Value = "红色个案数"
    .[D5].Value = "HRC17"
    .[E4].Value = "已经解决个案数"
    .[E5].Value = "HRC18"
    .[F4].Value = "解决百分比"
    .[F5].Value = "HRC19"
    .[G4].Value = "未解决个案数"
    .[G5].Value = "HRC20"
    .[H4].Value = "未解决百分比"
    .[H5].Value = "HRC21"
    .[I4].Value = "起始年"
    .[I5].Value = "HRC22"
    .[J4].Value = "起始月"
    .[J5].Value = "HRC23"
    .[K4].Value = "终止年"
    .[K5].Value = "HRC24"
    .[L4].Value = "终止月"
    .[L5].Value = "HRC25"
    .[M4].Value = "4个问题分类之一"
    .[M5].Value = "HRC26"
    .[N4].Value = "4个问题分类之二"
    .[N5].Value = "HRC27"
    .[O4].Value = "4个问题分类之三"
    .[O5].Value = "HRC28"
    .[P4].Value = "4个问题分类之四"
    .[P5].Value = "HRC29"
    .[B7].Value = "可能存在的第5个问题分类"
    .[B8].Value = "HRC30"
    .[C7].Value = "可能存在的第6个问题分类"
    .[C8].Value = "HRC31"
    .[D7].Value = "可能存在的第7个问题分类"
    .[D8].Value = "HRC32"
    .[E7].Value = "可能存在的第8个问题分类"
    .[E8].Value = "HRC33"
    
    .Columns("A:S").AutoFit
End With

IncludeFacCode = currentUseWB.Worksheets("Sheet1").[B2].Value
HRC00 = Trim(Left(IncludeFacCode, InStr(IncludeFacCode, "-") - 1))

IncludeStartYearAndMonth = currentUseWB.Worksheets("Sheet1").[A2].Value
HRC01 = Trim(Left(IncludeStartYearAndMonth, InStr(IncludeStartYearAndMonth, "年") - 1))


HRC02 = Mid(IncludeStartYearAndMonth, InStr(IncludeStartYearAndMonth, "年") + 1, InStr(IncludeStartYearAndMonth, "月") - InStr(IncludeStartYearAndMonth, "年") - 1)

IncludeEndYearAndMonth = currentUseWB.Worksheets("Sheet1").[A65536].End(xlUp).Value
HRC03 = Trim(Left(IncludeEndYearAndMonth, InStr(IncludeEndYearAndMonth, "年") - 1))

HRC04 = Mid(IncludeEndYearAndMonth, InStr(IncludeEndYearAndMonth, "年") + 1, InStr(IncludeEndYearAndMonth, "月") - InStr(IncludeEndYearAndMonth, "年") - 1)

HRC05 = HRC01
HRC06 = HRC02
HRC07 = HRC03
HRC08 = HRC04

HRC09 = currentUseWB.Worksheets("Sheet1").[A65536].End(xlUp).Row - 1

HRC10 = currentUseWB.Worksheets("SourceData").[A2].End(xlDown).Row - 1

ctGreen = 0
ctYellow = 0
ctRed = 0
ctZixun = 0
ctXinli = 0
ctTousu = 0

With currentUseWB.Worksheets("Sheet1")

For i = 2 To .[A65536].End(xlUp).Row
    If Trim(.Range("C" & i).Value) = "绿" Then ctGreen = ctGreen + 1
    If Trim(.Range("C" & i).Value) = "黄" Then ctYellow = ctYellow + 1
    If Trim(.Range("C" & i).Value) = "红" Then ctRed = ctRed + 1
    If Trim(.Range("D" & i).Value) = "咨询" Then ctZixun = ctZixun + 1
    If Trim(.Range("D" & i).Value) = "投诉" Then ctTousu = ctTousu + 1
    If Trim(.Range("D" & i).Value) = "心理" Then ctXinli = ctXinli + 1
Next

End With

HRC11 = ctZixun
HRC12 = Format(ctZixun / HRC09, "0.00%")

HRC13 = ctXinli
HRC14 = Format(ctXinli / HRC09, "0.00%")

HRC14Dot5 = ctTousu
HRC14Dot6 = Format(ctTousu / HRC09, "0.00%")

HRC15 = ctGreen
HRC16 = ctYellow
HRC17 = ctRed

HRC20 = InputBox(prompt:="请输入未解决个案数，如果个案已经全部解决，则此处输入0或者留空直接点击确定")

HRC20 = Val(HRC20)

HRC21 = Format(HRC20 / HRC09, "0.00%")

HRC18 = HRC09 - HRC20
HRC19 = Format(HRC18 / HRC09, "0.00%")

HRC22 = HRC01
HRC23 = HRC02
HRC24 = HRC03
HRC25 = HRC04

HRC26 = currentUseWB.Worksheets("SourceData").[A2].Value
HRC27 = currentUseWB.Worksheets("SourceData").[A3].Value
HRC28 = currentUseWB.Worksheets("SourceData").[A4].Value
HRC29 = currentUseWB.Worksheets("SourceData").[A5].Value
If currentUseWB.Worksheets("SourceData").[B5].Value = currentUseWB.Worksheets("SourceData").[B6].Value Then
    If currentUseWB.Worksheets("SourceData").[B6].Value = currentUseWB.Worksheets("SourceData").[B7].Value Then
        If currentUseWB.Worksheets("SourceData").[B7].Value = currentUseWB.Worksheets("SourceData").[B8].Value Then
            If currentUseWB.Worksheets("SourceData").[B8].Value = currentUseWB.Worksheets("SourceData").[B9].Value Then
                HRC30 = currentUseWB.Worksheets("SourceData").[A6].Value
                HRC31 = currentUseWB.Worksheets("SourceData").[A7].Value
                HRC32 = currentUseWB.Worksheets("SourceData").[A8].Value
                HRC33 = currentUseWB.Worksheets("SourceData").[A9].Value
            Else
                HRC30 = currentUseWB.Worksheets("SourceData").[B6].Value
                HRC31 = currentUseWB.Worksheets("SourceData").[A7].Value
                HRC32 = currentUseWB.Worksheets("SourceData").[A8].Value
            End If
        Else
            HRC30 = currentUseWB.Worksheets("SourceData").[A6].Value
            HRC31 = currentUseWB.Worksheets("SourceData").[A7].Value
        End If
    Else
        HRC30 = currentUseWB.Worksheets("SourceData").[A6].Value
    End If
End If


With currentUseWB.Worksheets("ToWordDoc")


    .[B3].Value = HRC00

    .[C3].Value = HRC01

    .[D3].Value = HRC02

    .[E3].Value = HRC03

    .[F3].Value = HRC04

    .[G3].Value = HRC05

    .[H3].Value = HRC06

    .[I3].Value = HRC07

    .[J3].Value = HRC08

    .[K3].Value = HRC09

    .[L3].Value = HRC10

    .[M3].Value = HRC11

    .[N3].Value = HRC12

    .[O3].Value = HRC13

    .[P3].Value = HRC14

    .[Q3].Value = HRC14Dot5

    .[R3].Value = HRC14Dot6

    .[B6].Value = HRC15

    .[C6].Value = HRC16

    .[D6].Value = HRC17

    .[E6].Value = HRC18

    .[F6].Value = HRC19

    .[G6].Value = HRC20

    .[H6].Value = HRC21

    .[I6].Value = HRC22

    .[J6].Value = HRC23

    .[K6].Value = HRC24

    .[L6].Value = HRC25

    .[M6].Value = HRC26

    .[N6].Value = HRC27

    .[O6].Value = HRC28

    .[P6].Value = HRC29

    .[B9].Value = HRC30

    .[C9].Value = HRC31

    .[D9].Value = HRC32

    .[E9].Value = HRC33
    
End With

'====================================

'将Excel中的三幅图保存到同一文件夹下
Set pie = currentUseWB.Worksheets.Add
pie.Name = "pie"
currentUseWB.Sheets("PieChart").ChartArea.Copy
currentUseWB.Sheets("pie").Paste

Set dgender = currentUseWB.Worksheets.Add
dgender.Name = "dgender"
currentUseWB.Sheets("Gender").ChartArea.Copy
currentUseWB.Sheets("dgender").Paste

Set dinmethod = currentUseWB.Worksheets.Add
dinmethod.Name = "dinmethod"
currentUseWB.Sheets("InMethod").ChartArea.Copy
currentUseWB.Sheets("dinmethod").Paste


currentUseWB.Worksheets("pie").ChartObjects(1).Chart.Export currentUseWB.Path & "\PieChart.png"
currentUseWB.Worksheets("dgender").Activate
ActiveSheet.ChartObjects(1).Chart.Export currentUseWB.Path & "\Gender.png"
currentUseWB.Worksheets("dinmethod").Activate
ActiveSheet.ChartObjects(1).Chart.Export currentUseWB.Path & "\InMethod.png"
'====================================


MsgBox "写入word文档所需数据处理完毕，点击确定开始生成Word文档，注意制定的阶段性报告模板【热线阶段性报告模版.docx】必须在此文件夹下！"

If Val(HRC30) = 0 Then
    endCellNumber = 30
ElseIf Val(HRC31) = 0 Then
    endCellNumber = 31
ElseIf Val(HRC32) = 0 Then
    endCellNumber = 32
ElseIf Val(HRC33) = 0 Then
    endCellNumber = 33
End If

Set WordObject = CreateObject("Word.Application")
currentPath = currentUseWB.Path

If Len(HRC02) = 1 Then HRC02 = "0" & HRC02
If Len(HRC04) = 1 Then HRC04 = "0" & HRC04

docName = HRC00 & " " & HRC01 & "-" & HRC03 & HRC04

HRC02 = Val(HRC02)
HRC04 = Val(HRC04)

FileCopy currentPath & "\热线阶段性报告模版.docx", currentPath & "\" & docName & ".docx"

'将所有HRC变量加入字典
Dim dict
Set dict = CreateObject("Scripting.Dictionary")

dict.Add "HRC00", HRC00
dict.Add "HRC01", HRC01
dict.Add "HRC02", HRC02
dict.Add "HRC03", HRC03
dict.Add "HRC04", HRC04
dict.Add "HRC05", HRC05
dict.Add "HRC06", HRC06
dict.Add "HRC07", HRC07
dict.Add "HRC08", HRC08
dict.Add "HRC09", HRC09
dict.Add "HRC10", HRC10
dict.Add "HRC11", HRC11
dict.Add "HRC12", HRC12
dict.Add "HRC13", HRC13
dict.Add "HRC14", HRC14
dict.Add "HRC15", HRC15
dict.Add "HRC16", HRC16
dict.Add "HRC17", HRC17
dict.Add "HRC18", HRC18
dict.Add "HRC19", HRC19
dict.Add "HRC20", HRC20
dict.Add "HRC21", HRC21
dict.Add "HRC22", HRC22
dict.Add "HRC23", HRC23
dict.Add "HRC24", HRC24
dict.Add "HRC25", HRC25
dict.Add "HRC26", HRC26
dict.Add "HRC27", HRC27
dict.Add "HRC28", HRC28
dict.Add "HRC29", HRC29
dict.Add "HRC30", HRC30
dict.Add "HRC31", HRC31
dict.Add "HRC32", HRC32
dict.Add "HRC33", HRC33
dict.Add "HRC14Dot5", HRC14Dot5
dict.Add "HRC14Dot6", HRC14Dot6

'=============
'写好word表格数据所需的框架
numOfAllCases = currentUseWB.Worksheets("Sheet1").[A65536].End(xlUp).Row - 1
realEndRow = currentUseWB.Worksheets("SourceData").[A2].End(xlDown).Row

  currentUseWB.Worksheets("SourceData").Range("C" & (realEndRow + 1)).Value = "咨询个案总数"
  currentUseWB.Worksheets("SourceData").Range("D" & (realEndRow + 1)).Value = ctZixun
  currentUseWB.Worksheets("SourceData").Range("E" & (realEndRow + 1)).Value = "占总个案数百分比"
   currentUseWB.Worksheets("SourceData").Range("F" & (realEndRow + 1)).Value = Format(ctZixun / numOfAllCases, "0.00%")

  currentUseWB.Worksheets("SourceData").Range("G" & (realEndRow + 1)).Value = "心理个案总数"
    currentUseWB.Worksheets("SourceData").Range("H" & (realEndRow + 1)).Value = ctXinli
            currentUseWB.Worksheets("SourceData").Range("I" & (realEndRow + 1)).Value = "占总个案数百分比"
        currentUseWB.Worksheets("SourceData").Range("J" & (realEndRow + 1)).Value = Format(ctXinli / numOfAllCases, "0.00%")
    
  currentUseWB.Worksheets("SourceData").Range("K" & (realEndRow + 1)).Value = "投诉个案总数"
    currentUseWB.Worksheets("SourceData").Range("L" & (realEndRow + 1)).Value = ctTousu
  currentUseWB.Worksheets("SourceData").Range("M" & (realEndRow + 1)).Value = "占总个案数百分比"
    currentUseWB.Worksheets("SourceData").Range("N" & (realEndRow + 1)).Value = Format(ctTousu / numOfAllCases, "0.00%")
  
'=============



With WordObject

 .Documents.Open (currentPath & "\" & docName & ".docx")
.Visible = False
    With .Selection.Find
        For i = 0 To 33
            If Len(i) = 1 Then
                si = "0" & i
            Else
                si = i
            End If
            siIndex = "HRC" & si
            .Text = siIndex
            .Replacement.Text = dict.Item(siIndex)
            .Execute Replace:=2 '全部替换
        Next
        
            .Text = "HRC40"
            .Replacement.Text = dict.Item("HRC14Dot5")
            .Execute Replace:=2 '全部替换
            .Text = "HRC41"
            .Replacement.Text = dict.Item("HRC14Dot6")
            .Execute Replace:=2 '全部替换
            
            
            
            
         Set hiDict = CreateObject("scripting.dictionary")
         endRow = currentUseWB.Worksheets("SourceData").Range("s2").Value
         ky = 1
         For i = 2 To endRow
            For j = 4 To 18 Step 2
                temp = currentUseWB.Worksheets("SourceData").Cells(i, j).Value
                If ky Mod 2 = 0 Then temp = Format(temp, "0.00%")
                hiDict.Add ky, temp
                ky = ky + 1
            Next
         Next
         
        For i = 1 To 40
            If Len(i) = 1 Then
                si = "0" & i
            Else
                si = i
            End If
            siIndex = "Four" & si
            .Text = siIndex
            .Replacement.Text = hiDict.Item(i)
            .Execute Replace:=2 '全部替换
        Next
    
    End With
    
    
Rem 下满是未完成的工作

'填充第四部分的文字部分

'填充第三和第六部分的图表


.Selection.Find.Text = "反映最多的个案问题分类"
.Selection.Find.Execute
.Selection.MoveRight 1
.Selection.InlineShapes.AddPicture Filename:= _
         currentPath & "\PieChart.png", _
         LinkToFile:=False, SaveWithDocument:=True

.Selection.Find.Text = "热线个案男女比例如下图"
.Selection.Find.Execute
.Selection.MoveRight 1
.Selection.InlineShapes.AddPicture Filename:= _
         currentPath & "\Gender.png", _
         LinkToFile:=False, SaveWithDocument:=True
         
         
.Selection.Find.Text = "热线个案接入方式比例："
.Selection.Find.Execute
.Selection.MoveRight 1
.Selection.InlineShapes.AddPicture Filename:= _
         currentPath & "\InMethod.png", _
         LinkToFile:=False, SaveWithDocument:=True
        
    

       
       
       '填入word文档最开始的表格
        With .ActiveDocument.Tables(1) 'Word表格
        
                '处理心理数据类型
                j = 8
                For i = 2 To realEndRow
                    If currentUseWB.Worksheets("SourceData").Cells(i, j).Value <> 0 Then
                       .Cell(2, 2) = currentUseWB.Worksheets("SourceData").Cells(i, 1).Value
                       .Cell(2, 3) = currentUseWB.Worksheets("SourceData").Cells(i, j).Value
                       .Cell(2, 4) = Format(currentUseWB.Worksheets("SourceData").Cells(i, j).Value / numOfAllCases, "0.00%")
                       .Cell(2, 5) = Format(currentUseWB.Worksheets("SourceData").Cells(realEndRow + 1, j + 2).Value, "0.00%") '此处可能显示不为百分比
                    End If
                Next
                
                '处理咨询数据类型
                j = 4
                m = 3
                For i = 2 To realEndRow
                    If currentUseWB.Worksheets("SourceData").Cells(i, j).Value <> 0 Then
                       .Cell(m, 2) = currentUseWB.Worksheets("SourceData").Cells(i, 1).Value
                       .Cell(m, 3) = currentUseWB.Worksheets("SourceData").Cells(i, j).Value
                       .Cell(m, 4) = Format(currentUseWB.Worksheets("SourceData").Cells(i, j).Value / numOfAllCases, "0.00%")
                       '.Cell(m, 5) = currentUseWB.Worksheets("SourceData").Cells(realEndRow, j + 2).Value
                       m = m + 1
                    End If
                Next
                 .Cell(3, 5) = Format(currentUseWB.Worksheets("SourceData").Cells(realEndRow + 1, j + 2).Value, "0.00%")
                 
                 '处理投诉事件类型
                 j = 12
                 m = 17
                For i = 2 To realEndRow
                    If currentUseWB.Worksheets("SourceData").Cells(i, j).Value <> 0 Then
                       .Cell(m, 2) = currentUseWB.Worksheets("SourceData").Cells(i, 1).Value
                       .Cell(m, 3) = currentUseWB.Worksheets("SourceData").Cells(i, j).Value
                       .Cell(m, 4) = Format(currentUseWB.Worksheets("SourceData").Cells(i, j).Value / numOfAllCases, "0.00%")
                       '.Cell(m, 5) = currentUseWB.Worksheets("SourceData").Cells(realEndRow, j + 2).Value
                       m = m + 1
                    End If
                Next
                
                .Cell(17, 5) = Format(currentUseWB.Worksheets("SourceData").Cells(realEndRow + 1, j + 2).Value, "0.00%")
            

        End With


End With

WordObject.Documents.Save

WordObject.Quit

Set WordObject = Nothing

'删除文件夹下的图片
Kill currentPath & "\PieChart.png"
Kill currentPath & "\Gender.png"
Kill currentPath & "\InMethod.png"

currentUseWB.Worksheets("Sheet1").Activate
currentUseWB.Worksheets("Sheet1").[A1].Select
MsgBox "生成完毕，请在本文件所在文件夹中找到阶段性报告的Word文档。"

End Sub

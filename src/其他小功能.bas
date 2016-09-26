Attribute VB_Name = "其他小功能"
'Date: 2016.8.1-2016.9.18
'Author: Vinson Wei
'Purpose: Aditional generic functions
Public Sub SumAllToOne(control As IRibbonControl)
Dim f$
Dim destinationSht, sht As Worksheet
Set currentUseWB = ActiveWorkbook
f = Dir(currentUseWB.Path & "\*.*")
If f = currentUseWB.Name Then f = Dir
If f = "" Then
    MsgBox "此文件夹下无其他文件。"
    Exit Sub
End If

For Each sht In currentUseWB.Worksheets
    If sht.Name = "Question Sheet" Then
        Set destinationSht = currentUseWB.Worksheets("Question Sheet")
        Exit For
    Else
        Set destinationSht = currentUseWB.Worksheets(1)
    End If
    
Next

On Error GoTo 0

Do
    If f <> currentUseWB.Name Then
        Set WB = Workbooks.Open(currentUseWB.Path & "\" & f)
        For Each sht In WB.Worksheets
            sht.Copy before:=destinationSht
        Next
        WB.Close False
    End If
    f = Dir
Loop Until f = ""
MsgBox "汇总完毕。"
End Sub

Sub NoResOpenCopyCloseOneiMonthlyReport(control As IRibbonControl)


Dim f$, WB As Workbook, pw$, loc$, sht As Worksheet, nameList

Set currentUseWB = ActiveWorkbook
On Error Resume Next
Application.DisplayAlerts = False
currentUseWB.Worksheets("Sheet2").Delete
currentUseWB.Worksheets("Sheet3").Delete
Application.DisplayAlerts = True
On Error GoTo 0
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
f = Dir(currentUseWB.Path & "\*.xls*")
Do

        '查询个案文件，打开月报文件

        If f <> currentUseWB.Name Then
        
        Set WB = Workbooks.Open(Filename:=currentUseWB.Path & "\" & f)
        
        
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
        WB.Worksheets(1).Range("A3").Copy _
        currentUseWB.Worksheets("Sheet1").Range("B" & iLInSheet1)
        WB.Worksheets(1).Range("C3").Copy _
        currentUseWB.Worksheets("Sheet1").Range("C" & iLInSheet1)
        EventAndQuestion = WB.Worksheets(1).Range("A5").Value
        midStart = InStr(EventAndQuestion, ":")
        If midStart = 0 Then midStart = InStr(EventAndQuestion, "：")
        If midStart = 0 Then
            MsgBox "在" & WB.Name & "文件的问题分类栏里面没有发现中文或者英文冒号，点击确定退出文件。"
            Exit Sub
        End If
        
        EventString = Trim(Mid(EventAndQuestion, 1, midStart - 1))
        QuestionString = Trim(Mid(EventAndQuestion, midStart + 1, 100))
        
        currentUseWB.Worksheets("Sheet1").Range("D" & iLInSheet1).Value = EventString
        currentUseWB.Worksheets("Sheet1").Range("E" & iLInSheet1).Value = QuestionString
                
        WB.Worksheets(1).Range("K3").Copy _
        currentUseWB.Worksheets("Sheet1").Range("F" & iLInSheet1)
                WB.Worksheets(1).Range("F3").Copy _
        currentUseWB.Worksheets("Sheet1").Range("G" & iLInSheet1)
                        
        WB.Worksheets(1).Range("B5").Copy _
        currentUseWB.Worksheets("Sheet1").Range("H" & iLInSheet1)
        
                WB.Worksheets(1).Range("B3").Copy _
        currentUseWB.Worksheets("Sheet1").Range("A" & iLInSheet1)
'
'        WB.Worksheets(1).Range("F3").Copy _
'        currentUseWB.Worksheets("Sheet1").Range("G" & iLInSheet1)
'        '从Case Sheet表中将沟通方式、事主性别搬运过来
'        WB.Worksheets(1).Range("F2", "G" & (nameList.Count + 1)).Copy _
'        currentUseWB.Worksheets("Sheet1").Range("F" & iLInSheet1)
'
        '写入年月
'        For i = iLInSheet1 To (iRInSheet1 - 1)
'        currentUseWB.Worksheets("Sheet1").Range("A" & i).Value = "'" & year & "年" & iMonth & "月"
'        iLInSheet1 = iLInSheet1 + 1
'        Next
        
        endCode = currentUseWB.Worksheets("Sheet1").Range("B" & (iLInSheet1 - 1)).Value
        
'        If Trim(endCode) <> nameList(nameList.Count - 1) Then
'            MsgBox "注意，现在发生了一个罕见的情况，我猜" & year & "年" & iMonth & "月的月报中" _
'            & "的某个个案工作表里面有两个个案（但共用一个个案编号），这" & "导致了导入时发生错位，点击确定，程序将尝试自动解决错位。"
'            GoTo iLNotEqualToiR
'        End If
        
        '正常的在这里开始
iLEqualToiR:
        iLInSheet1 = iRInSheet1
        
        WB.Close False
        
        End If
        
        f = Dir

Loop Until f = ""

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
MsgBox "导入完毕。"

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
currentUseWB.Worksheets("Sheet1").Range("A" & i).Value = "粗错啦！" '"'" &   "年" & iMonth & ""
iLInSheet1 = iLInSheet1 + 1
Next
MsgBox "Ok，错位问题已解决，点击继续程序运行。"
GoTo iLEqualToiR


End Sub
' Sub NoResCreatePivotTable(control As IRibbonControl)
' On Error Resume Next
'     Application.DisplayAlerts = False
'     currentUseWB.Sheets("PivotSheet").Delete
'     Application.DisplayAlerts = False
' On Error GoTo 0


' Dim pivotSheet As Worksheet, rg As Range, endCell As Range
' Set currentUseWB = ActiveWorkbook
' '创建数据源
' Set pivotSheet = currentUseWB.Worksheets.Add(after:=currentUseWB.Worksheets("Sheet1"))
' pivotSheet.Name = "PivotSheet"
' Set endCell = currentUseWB.Worksheets("Sheet1").Range("E2").End(xlDown)
' currentUseWB.Worksheets("Sheet1").Activate
' endCell.Select
' Set rg = Range(Range("D1"), endCell)

'  currentUseWB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'         rg, Version:=xlPivotTableVersion10).CreatePivotTable _
'         TableDestination:="PivotSheet!R3C1", TableName:="PivotTable", DefaultVersion:= _
'         xlPivotTableVersion10
'     Worksheets("PivotSheet").Select
'     currentUseWB.ShowPivotTableFieldList = True
'     With currentUseWB.Worksheets("PivotSheet").PivotTables("PivotTable").PivotFields("事件类型")
'         .Orientation = xlRowField
'         .Position = 1
'     End With
'         currentUseWB.Worksheets("PivotSheet").PivotTables("PivotTable").AddDataField currentUseWB.Worksheets("PivotSheet").PivotTables("PivotTable" _
'         ).PivotFields("问题分类"), "计数项：问题分类", xlCount
'     With currentUseWB.Worksheets("PivotSheet").PivotTables("PivotTable").PivotFields("问题分类")
'         .Orientation = xlColumnField
'         .Position = 1
'     End With

'     currentUseWB.ShowPivotTableFieldList = False
    
'     MsgBox "数据透视表生成完毕"
' End Sub

 Sub NoResCreatePieChart(control As IRibbonControl)
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
    If ctZixun = 0 Then GoTo Step1
    If j = 6 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, j - 2).Value / ctZixun, "0.00%")
Back1:
    If ctXinli = 0 Then GoTo Step2
    If j = 10 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, j - 2).Value / ctXinli, "0.00%")
Back2:
    If ctTousu = 0 Then GoTo Step3
    If j = 14 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = Format(currentUseWB.Worksheets("SourceData").Cells(i, j - 2).Value / ctTousu, "0.00%")
Back3:
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
    
'=====================================================

MsgBox "成功生成了一幅饼图和两幅柱状图，请注意如果两幅柱状图中数据只有一列（比如只有男或者只有QQ接入），则柱状图的格式可能有问题，需要手工修改格式"

Exit Sub

Step1:
   If j = 6 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = 0
 GoTo Back1
Step2:
    If j = 10 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = 0
 GoTo Back2
Step3:
    If j = 14 Then currentUseWB.Worksheets("SourceData").Cells(i, j).Value = 0
GoTo Back3

End Sub

Attribute VB_Name = "月报自动制作脚本v1"
'Date: 2016.8.1-2016.9.18
'Author: Vinson Wei
'Purpose: Automate the process of making a monthly report

Dim numOldStart, numOldEnd As String
Dim numNewStart, numNewEnd As String
Dim factory As String
Dim currentUseWB
Dim status As Integer
Dim i, lenOfi As Integer
Dim ni As Integer
Dim sIndex As String
Dim WB As Workbook
Dim CaseSheet, QuestionSheet, QuestionSheetInVertical, PerspectiveView As Worksheet


Sub ExecuteAllStepsAtOnce(control As IRibbonControl)

Application.DisplayAlerts = False

status = 0
Call Initialization(status)

If status <> 0 Then
    Select Case status
        Case 1
            MsgBox "未输入工厂代码，点击确定退出月报制作"
        Case 2
            MsgBox "未输入上月最小个案编号，点击确定退出月报制作"
        Case 3
            MsgBox "未输入上月最大个案编号，点击确定退出月报制作"
        Case 4
            MsgBox "未输入本月最小个案编号，点击确定退出月报制作"
        Case 5
            MsgBox "未输入本月最大个案编号，点击确定退出月报制作"
    End Select
Exit Sub
End If

DeleteOldSheets
InsertNewSheets
UpdateCaseSheet
UpdateQuestionSheet
UpdateQuestionSheetInVertical
UpdatePerspectiveView
LocateALLCursorsOnA1

Application.DisplayAlerts = True

Delay (1)


MsgBox "程序已经处理完毕，完成的任务有：删除上月所有个案工作表，插入本月所有个案工作表，更新Case Sheet" _
& "表中所有列的信息、更新Question Sheet表中所有列信息、更新Question Sheet in Vertical表中所有信息、重新选择数据透视表" _
& "数据源并刷新数据透视表，所有工作表光标都已经定位在A1单元格。【【注意【所有】这些更改都没有保存，如果执行的过程中有错误输入，请直接关闭文件，点击不保存即可，然后利用此文件重新开始月报制作】】，如有遗漏事项没有处理，请手动继续完成，最后手动保存"
End Sub
Sub Initialization(ByRef status As Integer)

factory = InputBox(prompt:="输入工厂代码", Title:="工厂代码")          '工厂代码

If factory = "" Then
status = 1
Exit Sub
End If

numOldStart = InputBox(prompt:="输入上月个案最小编号", Title:="输入上月个案最小编号")         '上月该工厂个案最小编号

If numOldStart = "" Then
status = 2
Exit Sub
End If

numOldEnd = InputBox(prompt:="输入上月个案最大编号", Title:="输入上月个案最大编号")        '上月该工厂个案最大编号

If numOldEnd = "" Then
status = 3
Exit Sub
End If

numNewStart = InputBox(prompt:="输入本月个案最小编号", Title:="输入本月个案最小编号")       '本月该工厂个案最小编号

If numNewStart = "" Then
status = 4
Exit Sub
End If

numNewEnd = InputBox(prompt:="输入本月个案最大编号", Title:="输入本月个案最大编号")       '本月该工厂个案最大编号

If numNewEnd = "" Then
status = 5
Exit Sub
End If

Set CaseSheet = Worksheets("Case Sheet")
Set QuestionSheet = Worksheets("Question Sheet")
Set QuestionSheetInVertical = Worksheets("Question Sheet in Vertical")
Set PerspectiveView = Worksheets("Perspective View")
Set currentUseWB = ActiveWorkbook
End Sub


Sub Delay(T As Single)
    Dim time1 As Single
    time1 = Timer
    Do
        DoEvents
    Loop While Timer - time1 < T
End Sub
Sub DeleteOldSheets() '删除上月所有个案工作表

For i = Val(numOldStart) To Val(numOldEnd)
lenOfi = Len(CStr(i))
If lenOfi = 1 Then
    sIndex = factory & "-" & "00" & CStr(i)
    'ActiveSheet.[A6] = sIndex
ElseIf lenOfi = 2 Then
    sIndex = factory & "-" & "0" & CStr(i)
Else
    sIndex = factory & "-" & CStr(i)
End If
On Error Resume Next
Worksheets(sIndex).Delete
Next

End Sub
Sub InsertNewSheets() '在月报中插入本月所有个案

Application.ScreenUpdating = False


For ni = Val(numNewStart) To Val(numNewEnd)
        lenOfni = Len(CStr(ni))
        If lenOfni = 1 Then
            sIndex = factory & "-" & "00" & CStr(ni)
        ElseIf lenOfni = 2 Then
            sIndex = factory & "-" & "0" & CStr(ni)
        Else
            sIndex = factory & "-" & CStr(ni)
        End If
        f = Dir(currentUseWB.Path & "\" & "??" & sIndex & "*" & ".xls")
        If f = "" Then f = Dir(currentUseWB.Path & "\" & "??" & sIndex & "*" & ".xlsx")
        If f = "" Then
            MsgBox "没有找到" & sIndex & "个案所在Excel工作簿，请检查同一文件夹下是否存在" & sIndex & "个案文件，如果没有此文件，" _
            & "从NAS复制对应的个案到同一文件夹下，如果该个案文件名格式和其他个案不一样，" _
            & "请修改为与其他个案相同的文件名格式，然后关闭程序，关闭此Excel文件，选择不保存更改，然后重新制作月报。"
            Exit Sub
        End If
        Workbooks.Open (currentUseWB.Path & "\" & f)
        Set WB = ActiveWorkbook
        
        If WB.Worksheets.Count > 1 Then
            MsgBox WB.Name & "个案工作簿里面有1张以上工作表，请前往查看改正，点击确定退出程序。"
            Exit Sub
        End If
        
  
        If WB.Worksheets(1).Name <> sIndex Then
            MsgBox "嗯...我猜" & sIndex & "个案文件中的工作表标签是错的，这个" & vbLf & "错误好像经常出现(●￣(?)￣●)，点击确定，程序自动尝试将其改成正确的工作表标签。"
            WB.Worksheets(1).Name = sIndex
            MsgBox "OK，个案文件中错误的工作表标签已经更改，点击确定继续程序运行。"
            
        End If
        WB.Worksheets(sIndex).Copy before:=currentUseWB.Worksheets("Question Sheet")
        If Trim(currentUseWB.Worksheets(sIndex).[E3].Value) <> factory Then
            MsgBox "哦..." & sIndex & "个案文件里面的工厂代码（Name of the Factory）即E3单元格里的内容不正确，点击确定后程序将自动更改【月报】中的工厂代码和【个案文件】中的工厂代码。"
            currentUseWB.Worksheets(sIndex).[E3].Value = factory
            WB.Worksheets(sIndex).[E3].Value = factory
            MsgBox "OK，【月报】和【个案文件】中的错误工厂代码（Name of the Factory）已经更改，点击确定继续程序运行。"
        End If
        
        endBrand = InStr(3, currentUseWB.Name, " ")
        If endBrand = 0 Then endBrand = InStr(3, currentUseWB.Name, " ")
        brand = Trim(Mid(currentUseWB.Name, 1, endBrand - 1))
        
        If Trim(currentUseWB.Worksheets(sIndex).[D3].Value) <> brand Then
            MsgBox "啊.." & sIndex & "个案文件里面的品牌代码（Brand）即D3单元格的内容不正确，点击确定后程序将自动更改【月报】中的品牌代码和【个案文件】中的品牌代码。"
            currentUseWB.Worksheets(sIndex).[D3].Value = brand
            WB.Worksheets(sIndex).[D3].Value = brand
            MsgBox "OK，【月报】和【个案文件】中的错误品牌代码（Brand）已经更改，点击确定继续程序运行。"
        End If
        
        If Trim(currentUseWB.Worksheets(sIndex).[A3].Value) <> sIndex Then
            MsgBox "程序发现" & sIndex & "个案文件里的个案编号（Serial Number）及A3单元格的内容有错误，点击确定后程序将自动更改【月报】中的个案编号和【个案文件】中的个案编号。"
            currentUseWB.Worksheets(sIndex).[A3].Value = sIndex
            WB.Worksheets(sIndex).[A3].Value = sIndex
            MsgBox "OK，【月报】和【个案文件】中的错误个案编号（Serial Number）已经更改，点击确定继续程序运行。"
        End If
        
        WB.Close True
        


Next


Application.ScreenUpdating = True


End Sub

Sub UpdateCaseSheet() '更新Case Sheet表中所有信息
Dim iInCaseSheet, ni, lenOfni As Integer, sIndex, toSheetNum As String
Dim newSpan, oldSpan As Integer

newSpan = Val(numNewEnd) - Val(numNewStart)
oldSpan = Val(numOldEnd) - Val(numOldStart)


Rem I.1 调整Case Sheet的数量栏，分别考虑本月个案比上月少/多/相等的情况
If (oldSpan) > (newSpan) Then
        startClearRow = newSpan + 3
        CaseSheet.Range(CStr(startClearRow) & ":65536").Clear
End If
For iInCaseSheet = 1 To (newSpan + 1)
Length = Len(CStr(iInCaseSheet))
        If Length = 1 Then
            sIndex = "00" & CStr(iInCaseSheet)
        ElseIf Length = 2 Then
            sIndex = "0" & CStr(iInCaseSheet)
        Else
            sIndex = CStr(iInCaseSheet)
        End If
        CaseSheet.Range("A" & CStr(iInCaseSheet + 1)).Value = "'" & sIndex
Next
'用格式刷将新增的序号刷成原来的格式
CaseSheet.[A2].Copy
CaseSheet.Range("A2", "A" & CStr(newSpan + 2)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False

Rem I.2 更新'数量'列的超链接
'(1)获得每次要更新的toSheet的名字toIndex
For iInCaseSheet = 1 To (newSpan + 1)
toSheetNum = Val(numNewStart) + iInCaseSheet - 1
lenOftoSheetNum = Len(CStr(toSheetNum))
If lenOftoSheetNum = 1 Then
    toNum = "00" & CStr(toSheetNum)
    toIndex = factory & "-" & toNum
ElseIf lenOftoSheetNum = 2 Then
    toNum = "0" & CStr(toSheetNum)
    toIndex = factory & "-" & toNum
Else
    toNum = CStr(toSheetNum)
    toIndex = factory & "-" & toNum
End If

'(2)处理数量栏显示的索引,获取每次要更新的单元格的值
lengOfIndex = Len(CStr(iInCaseSheet))
If lengOfIndex = 1 Then
toNumShow = "00" & iInCaseSheet
ElseIf lengOfIndex = 2 Then
toNumShow = "0" & iInCaseSheet
Else
toNumShow = iInCaseSheet
End If

'(3)有了上面两步的准备就可以更新超链接了
CaseSheet.Hyperlinks.Add Anchor:=CaseSheet.Range("A" & (iInCaseSheet + 1)), _
Address:="", _
SubAddress:="'" & toIndex & "'!A1", _
TextToDisplay:="'" & toNumShow
Next
'数量列的调整、超链接更新完毕

Rem II更新BAS分色列、工厂代号、个案编号、反映时间、沟通方式、事主性别
For iInCaseSheet = 1 To (newSpan + 1)
toSheetNum = Val(numNewStart) + iInCaseSheet - 1
lenOftoSheetNum = Len(CStr(toSheetNum))
If lenOftoSheetNum = 1 Then
    toNum = "00" & CStr(toSheetNum)
    toIndex = factory & "-" & toNum
ElseIf lenOftoSheetNum = 2 Then
    toNum = "0" & CStr(toSheetNum)
    toIndex = factory & "-" & toNum
Else
    toNum = CStr(toSheetNum)
    toIndex = factory & "-" & toNum
End If

'更新BAS分色
'更新文字内容
CaseSheet.Range("B" & CStr(iInCaseSheet + 1)).Value = Worksheets(toIndex).Range("C3").Value
'更新字体颜色
CaseSheet.Range("B" & CStr(iInCaseSheet + 1)).Font.Color = Worksheets(toIndex).Range("C3").Interior.Color

'工厂代码
CaseSheet.Range("C" & CStr(iInCaseSheet + 1)).Value = factory

'个案编号
CaseSheet.Range("D" & CStr(iInCaseSheet + 1)).Value = "'" & toNum

'反映时间
CaseSheet.Range("E" & CStr(iInCaseSheet + 1)).Value = Worksheets(toIndex).Range("B3").Value

'沟通方式
CaseSheet.Range("F" & CStr(iInCaseSheet + 1)).Value = Worksheets(toIndex).Range("K3").Value

'事主性别
CaseSheet.Range("G" & CStr(iInCaseSheet + 1)).Value = Worksheets(toIndex).Range("F3").Value
Next

'将工厂代码、个案编号、反映时间、事主性别刷成原来的格式
'主要用于防止newSpan>oldSpan的情况下多出来的行里的数据没有合适的格式
'BAS分色列比较特殊，暂不处理
CaseSheet.Range("C2").Copy

CaseSheet.Range("C2", "C" & CStr(newSpan + 2)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False

CaseSheet.Range("D2").Copy
CaseSheet.Range("D2", "D" & CStr(newSpan + 2)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False


CaseSheet.Range("E2").Copy
CaseSheet.Range("E2", "E" & CStr(newSpan + 2)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False


CaseSheet.Range("F2").Copy
CaseSheet.Range("F2", "F" & CStr(newSpan + 2)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False

CaseSheet.Range("G2").Copy
CaseSheet.Range("G2", "G" & CStr(newSpan + 2)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False

'处理BAS分色列的字体
If (CaseSheet.Range("B2").Value Like "*e*") Then
    With CaseSheet.Range("B2", "B" & CStr(newSpan + 2))
        .Font.Name = "Arial"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
Else
    With CaseSheet.Range("B2", "B" & CStr(newSpan + 2))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Name = "宋体"
        .Font.Size = 10
    End With
End If


Rem BAS分色列处理
 CaseSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlDiagonalDown).LineStyle = xlNone
    CaseSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlDiagonalUp).LineStyle = xlNone
    With CaseSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With CaseSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With CaseSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With CaseSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With CaseSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With CaseSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
Rem BAS分色列处理

End Sub

Sub UpdateQuestionSheet()
'更新QuesttionSheet
Dim newSpan, oldSpan As Integer

newSpan = Val(numNewEnd) - Val(numNewStart)
oldSpan = Val(numOldEnd) - Val(numOldStart)

Rem 更新个案编号列
If (oldSpan) > (newSpan) Then
        startClearRow = newSpan + 3
        QuestionSheet.Range(CStr(startClearRow) & ":65536").Clear
End If
For iInCaseSheet = 1 To (newSpan + 1)
                toSheetNum = Val(numNewStart) + iInCaseSheet - 1
                lenOftoSheetNum = Len(CStr(toSheetNum))
                If lenOftoSheetNum = 1 Then
                    toNum = "00" & CStr(toSheetNum)
                    toIndex = factory & "-" & toNum
                ElseIf lenOftoSheetNum = 2 Then
                    toNum = "0" & CStr(toSheetNum)
                    toIndex = factory & "-" & toNum
                Else
                    toNum = CStr(toSheetNum)
                    toIndex = factory & "-" & toNum
                End If
                QuestionSheet.Range("A" & CStr(iInCaseSheet + 1)).Value = toIndex
Next
'用格式刷将新增的序号刷成原来的格式
QuestionSheet.[A2].Copy
QuestionSheet.Range("A2", "A" & CStr(newSpan + 2)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False

Rem QuestionSheet表大循环开始
Rem QuestionSheet表大循环开始
Rem QuestionSheet表大循环开始
For iInQuestionSheet = 1 To (newSpan + 1)
toSheetNum = Val(numNewStart) + iInQuestionSheet - 1
lenOftoSheetNum = Len(CStr(toSheetNum))
If lenOftoSheetNum = 1 Then
    toNum = "00" & CStr(toSheetNum)
    toIndex = factory & "-" & toNum
ElseIf lenOftoSheetNum = 2 Then
    toNum = "0" & CStr(toSheetNum)
    toIndex = factory & "-" & toNum
Else
    toNum = CStr(toSheetNum)
    toIndex = factory & "-" & toNum
End If

'更新颜色
'更新文字内容
QuestionSheet.Range("B" & CStr(iInQuestionSheet + 1)).Value = Worksheets(toIndex).Range("C3").Value
'更新字体颜色
QuestionSheet.Range("B" & CStr(iInQuestionSheet + 1)).Font.Color = Worksheets(toIndex).Range("C3").Interior.Color

Rem 更新事件类型和问题分类

'midStart = WorksheetFunction.IfError(WorksheetFunction.Find("：", Worksheets(toIndex).Range("A5")), WorksheetFunction.Find(":", Worksheets(toIndex).Range("A5")))
midStart = InStr(Worksheets(toIndex).Range("A5"), ":")
If midStart = 0 Then midStart = InStr(Worksheets(toIndex).Range("A5"), "：")

If midStart <> 0 Then GoTo noError '如果成功检测到冒号，则跳过下面一大段错误处理代码

Rem 开始问题分类栏的错误处理代码
Rem 开始问题分类栏的错误处理代码
If midStart = 0 Then
MsgBox ("没有在" & toIndex & "个案工作表中的问题分类栏里找到冒号->：<-，请点击确定，程序将自动尝试将其他误输入的标点符号更改为冒号" _
& "或者添加中文或者英文冒号")

flag = 0

'很多个小if单元开始
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), ",")
        flag = 1
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "，")
        flag = 2
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), ".")
        flag = 3
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "。")
        flag = 4
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), ";")
        flag = 5
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "；")
        flag = 6
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "'")
        flag = 7
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "’")
        flag = 8
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "“")
        flag = 9
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), Chr(34))
        flag = 10
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "<")
        flag = 11
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "《")
        flag = 12
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), ">")
        flag = 13
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "》")
        flag = 14
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "?")
        flag = 15
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "？")
        flag = 16
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "|")
        flag = 17
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "[")
        flag = 18
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "【")
        flag = 19
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "]")
        flag = 20
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "】")
        flag = 21
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "、")
        flag = 22
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "/")
        flag = 23
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "\")
        flag = 24
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "”")
        flag = 25
    End If
    If midStart = 0 Then
        midStart = InStr(Worksheets(toIndex).Range("A5"), "’")
        flag = 25
    End If
'很多个小if单元结束

End If

If midStart = 0 Then
        MsgBox ("非常抱歉，在个案工作表" & toIndex & "的问题分类栏里没有匹配到下列可能误输入的标点符号：" _
        & ",，.。;；'’'“<《>》" & "?" & "？|[【]】、/\”" & Chr(34) _
        & "点击确定，程序会继续尝试直接添加中文或者英文冒号")
        
        '不存在标点符号，直接添加
        If (Worksheets(toIndex).Range("A5").Value Like "*co*") Or (Worksheets(toIndex).Range("A5").Value Like "*Co*") Then
            enGLetterLocation = InStr(Worksheets(toIndex).Range("A5").Value, "g")
            enLeft = Mid(Worksheets(toIndex).Range("A5").Value, 1, enGLetterLocation)
            enRight = Mid(Worksheets(toIndex).Range("A5").Value, enGLetterLocation + 1)
            Worksheets(toIndex).Range("A5").Value = enLeft & ":" & enRight
            MsgBox "在个案工作表" & toIndex & "的问题分类栏的合适位置成功添加了英文冒号，点击确定继续运行程序处理为完成的任务"
            midStart = InStr(Worksheets(toIndex).Range("A5").Value, ":")
            GoTo noError
        ElseIf (Worksheets(toIndex).Range("A5").Value Like "*投诉*") _
                Or (Worksheets(toIndex).Range("A5").Value Like "*咨询*") _
                Or (Worksheets(toIndex).Range("A5").Value Like "*心理*") Then
            boolTouSuExist = Worksheets(toIndex).Range("A5").Value Like "*投诉*"
            boolZiXunExist = Worksheets(toIndex).Range("A5").Value Like "*咨询*"
            boolXinLiExist = Worksheets(toIndex).Range("A5").Value Like "*心理*"
            If boolTouSuExist Then
                zhKeyHanzi = InStr(Worksheets(toIndex).Range("A5").Value, "诉")
                zhLeft = Mid(Worksheets(toIndex).Range("A5").Value, 1, zhKeyHanzi)
                zhLeft = Mid(Worksheets(toIndex).Range("A5").Value, zhKeyHanzi + 1)
                Worksheets(toIndex).Range("A5").Value = zhLeft & "：" & zhRight
            ElseIf boolZiXunExist Then
                zhKeyHanzi = InStr(Worksheets(toIndex).Range("A5").Value, "询")
                zhLeft = Mid(Worksheets(toIndex).Range("A5").Value, 1, zhKeyHanzi)
                zhLeft = Mid(Worksheets(toIndex).Range("A5").Value, zhKeyHanzi + 1)
                Worksheets(toIndex).Range("A5").Value = zhLeft & "：" & zhRight
            ElseIf boolXinLiExist Then
                zhKeyHanzi = InStr(Worksheets(toIndex).Range("A5").Value, "理")
                zhLeft = Mid(Worksheets(toIndex).Range("A5").Value, 1, zhKeyHanzi)
                zhLeft = Mid(Worksheets(toIndex).Range("A5").Value, zhKeyHanzi + 1)
                Worksheets(toIndex).Range("A5").Value = zhLeft & "：" & zhRight
            End If
            midStart = InStr(Worksheets(toIndex).Range("A5").Value, "：")
            GoTo noError
        Else
            MsgBox "没有在" & toIndex & "工作表的问题分类(Problem)里找到中文或者英文的事件类型，请【关闭Excel程序】，抛弃此工作文件夹，从" _
            & "NAS上重新下载个案和上个月的月报模板，修改" & toIndex & "工作表的错误以及其他个案工作簿中可能的错误以后，再重新导入模块，按照要求重新运行程序"
        End If
        '添加完毕
ElseIf midStart <> 0 Then '找到了误输入的标点符号，再注意看一下这段代码，与前面矛盾
    errChar = ""
    boolEn = (Worksheets(toIndex).Range("A5").Value Like "*co*") Or (Worksheets(toIndex).Range("A5").Value Like "*Co*")
    If Not (boolEn) Then
        Select Case flag
            Case 1
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, ",", "：")
                errChar = ","
            Case 2
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "，", "：")
                errChar = "，"
            Case 3
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, ".", "：")
                errChar = "."
            Case 4
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "。", "：")
                errChar = "。"
            Case 5
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, ";", "：")
                errChar = ";"
            Case 6
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "；", "：")
                errChar = "；"
            Case 7
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "'", "：")
                errChar = "'"
            Case 8
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "‘", "：")
                errChar = "‘"
            Case 9
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "“", "：")
                errChar = "“"
            Case 10
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, Chr(34), "：")
                errChar = Chr(34)
            Case 11
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "<", "：")
                errChar = "<"
            Case 12
            Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "《", "：")
                errChar = "《"
            Case 13
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, ">", "：")
                errChar = ">"
            Case 14
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "》", "：")
                errChar = "》"
            Case 15
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "?", "：")
                errChar = "?"
            Case 16
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "？", "：")
                errChar = "？"
            Case 17
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "|", "：")
                errChar = "|"
            Case 18
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "[", "：")
                errChar = "["
            Case 19
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "【", "：")
                errChar = "【"
            Case 20
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "]", "：")
                errChar = "]"
            Case 21
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "】", "：")
            errChar = "】"
                Case 22
            Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "、", "：")
                errChar = "、"
            Case 23
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "/", "：")
                errChar = "/"
            Case 24
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "\", "：")
                errChar = "\"
            Case 25
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "”", "：")
                errChar = "”"
            Case 26
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "’", "：")
                errChar = "’"
        End Select
    Else
        Select Case flag
            Case 1
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, ",", ":")
                errChar = ","
            Case 2
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "，", ":")
                errChar = "，"
            Case 3
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, ".", ":")
                errChar = "."
            Case 4
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "。", ":")
                errChar = "。"
            Case 5
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, ";", ":")
                errChar = ";"
            Case 6
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "；", ":")
                errChar = "；"
            Case 7
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "'", ":")
                errChar = "'"
            Case 8
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "‘", ":")
                errChar = "‘"
            Case 9
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "“", ":")
                errChar = "“"
            Case 10
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, Chr(34), ":")
                errChar = Chr(34)
            Case 11
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "<", ":")
                errChar = "<"
            Case 12
            Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "《", ":")
                errChar = "《"
            Case 13
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, ">", ":")
                errChar = ">"
            Case 14
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "》", ":")
                errChar = "》"
            Case 15
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "?", ":")
                errChar = "?"
            Case 16
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "？", ":")
                errChar = "？"
            Case 17
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "|", ":")
                errChar = "|"
            Case 18
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "[", ":")
                errChar = "["
            Case 19
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "【", ":")
                errChar = "【"
            Case 20
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "]", ":")
                errChar = "]"
            Case 21
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "】", ":")
            errChar = "】"
                Case 22
            Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "、", ":")
                errChar = "、"
            Case 23
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "/", ":")
                errChar = "/"
            Case 24
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "\", ":")
                errChar = "\"
            Case 25
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "”", ":")
                errChar = "”"
            Case 26
                Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, "’", ":")
                errChar = "’"
        End Select
    End If

    maoHao = iif(boolEn, "英文冒号", "中文冒号")
    MsgBox "程序成功检测到" & toIndex & "工作表中的问题分类（Problem Type）里面误输入了" & Chr(34) & errChar & Chr(34) _
    & "点击确定程序将自动将该符号改为" & maoHao
    Worksheets(toIndex).Range("A5").Value = Replace(Worksheets(toIndex).Range("A5").Value, errChar, "：")
    MsgBox "程序成功将误输入标点符号替换成了，" & maoHao & "点击确定继续程序的进行"
    If boolEn Then
        midStart = InStr(Worksheets(toIndex).Range("A5").Value, ":")
    Else
        midStart = InStr(Worksheets(toIndex).Range("A5").Value, "：")
    End If
    GoTo noError

End If

Rem 结束问题分类栏的错误处理代码
Rem 结束问题分类栏的错误处理代码



noError:
QuestionSheet.Range("C" & CStr(iInQuestionSheet + 1)).Value = Trim(Mid(Worksheets(toIndex).Range("A5").Value, 1, midStart - 1)) '这个大循环主要就是填写Question Sheet表中的C列的事件类型和D列的问题分类，本行填充C列
QuestionSheet.Range("D" & CStr(iInQuestionSheet + 1)).Value = Trim(Mid(Worksheets(toIndex).Range("A5"), midStart + 1, 100)) '这个大循环主要就是填写Question Sheet表中的C列的事件类型和D列的问题分类，本行填充D列
Next

Rem QuestionSheet表大循环结束
Rem QuestionSheet表大循环结束
Rem QuestionSheet表大循环结束



'解决其他列可能出现的格式问题，颜色列较特殊，除外。
QuestionSheet.[C2].Copy
QuestionSheet.Range("C2", "C" & CStr(newSpan + 2)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False

QuestionSheet.[d2].Copy
QuestionSheet.Range("D2", "D" & CStr(newSpan + 2)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False

'设置颜色列的字体
If (QuestionSheet.Range("B2").Value Like "*e*") Then
    With QuestionSheet.Range("B2", "B" & CStr(newSpan + 2))
        .Font.Name = "Arial"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
Else
    With QuestionSheet.Range("B2", "B" & CStr(newSpan + 2))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Name = "宋体"
        .Font.Size = 10
    End With
End If

Rem 下面这段代码仅仅是为了设置下框线
 QuestionSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlDiagonalDown).LineStyle = xlNone
    QuestionSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlDiagonalUp).LineStyle = xlNone
    With QuestionSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With QuestionSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With QuestionSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With QuestionSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With QuestionSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With QuestionSheet.Range("B2", "B" & CStr(newSpan + 2)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
Rem 上面这段代码仅仅是为了设置下框线


End Sub

Sub UpdateQuestionSheetInVertical()
Dim newSpan, oldSpan As Integer

newSpan = Val(numNewEnd) - Val(numNewStart)
oldSpan = Val(numOldEnd) - Val(numOldStart)

Rem 更新个案编号列
If (oldSpan) > (newSpan) Then
        startClearRow = newSpan + 3
        QuestionSheetInVertical.Range(CStr(startClearRow) & ":65536").Clear
End If
QuestionSheet.Range("C1" & ":D" & (newSpan + 2)).Copy QuestionSheetInVertical.Range("A1") '是newSpan+2
End Sub

Sub UpdatePerspectiveView()
    
    '选择数据透视表
    PerspectiveView.PivotTables("数据透视表1").PivotSelect "", xlDataAndLabel, True
    '更改数据源
    PerspectiveView.PivotTables("数据透视表1").ChangePivotCache ActiveWorkbook.PivotCaches. _
        Create(SourceType:=xlDatabase, SourceData:= _
        QuestionSheetInVertical.[A1].CurrentRegion, _
        Version:=xlPivotTableVersion10)
    '两次刷新数据表
    PerspectiveView.PivotTables("数据透视表1").PivotCache.Refresh
    PerspectiveView.PivotTables("数据透视表1").PivotCache.Refresh
    'PerspectiveView.Range("A1").Select
End Sub

Sub LocateALLCursorsOnA1()
Dim sht As Worksheet
For Each sht In currentUseWB.Worksheets
sht.Activate
sht.[A1].Select
Next
CaseSheet.Activate
CaseSheet.[A1].Select
End Sub

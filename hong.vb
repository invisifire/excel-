
Sub mijie()
'
' mijie Macro


'
    Sheets("确认密接库-先不录").Columns("A:Z").Delete Shift:=xlShiftToLeft
    '选择表单’确认密接库‘的第A到Z列。删除，删除规则是向左补齐
    Sheets("潜在密接全库").Columns("A:Z").Copy Sheets("确认密接库-先不录").[A1]
    '选择表单‘潜在密接库’的第A到Z列复制到‘确认密接库’a1单元格
    Sheets("确认密接库-先不录").Range("A1").AutoFilter
    '选择表单‘确认密接库’的A1单元格，设置为自动筛选
    With Sheets("确认密接库-先不录").AutoFilter.Sort
        With .SortFields
            .Clear
            .Add Key:=Range("B1"), Order:=xlDescending
        End With
        .Header = xlYes
        .MatchCase = False
        .SortMethod = xlPinYin
        .Orientation = xlSortColumns
        .Apply
    End With
    '上面全部是选择筛选B1降序（日期）
    Application.CutCopyMode = False
    Sheets("确认密接库-先不录").Columns("A:Z").RemoveDuplicates Columns:=3, Header:=xlYes
    'Columns:=3是姓名列号  此句为删除第3列的重复项保留首项
    Sheets("确认密接库-先不录").Columns("A:Z").AutoFilter Field:=18, Criteria1:=Array("阳性"), Operator:=xlFilterValues
    'Field:=18是病例复核结果 此句为选择18列筛选阳性
    ActiveWorkbook.Names.Add Name:="'确认密接库-先不录'!_FilterDatabase", RefersTo:="='确认密接库-先不录'!$A:$Z", Visible:=False
    'Range("A1").Select 此句为将筛选结果保存更新到表带
End Sub

Sub Macro1()
'
    Sheets("已确认阳性库").Columns("A:AM").Delete Shift:=xlShiftToLeft
    Sheets("初筛阳性库").Columns("A:AM").Copy Sheets("已确认阳性库").[A1]
    Sheets("已确认阳性库").Range("A1").AutoFilter
    With Sheets("已确认阳性库").AutoFilter.Sort
        With .SortFields
            .Clear
            .Add Key:=Range("V1"), Order:=xlDescending' Range("V1") 初筛阳性日期
        End With
        .Header = xlYes
        .MatchCase = False
        .SortMethod = xlPinYin
        .Orientation = xlSortColumns
        .Apply
    End With
    Application.CutCopyMode = False
    Sheets("已确认阳性库").Columns("A:AM").RemoveDuplicates Columns:=2, Header:=xlYes
    'Columns:=2是姓名列号
    Sheets("已确认阳性库").Columns("A:AM").AutoFilter Field:=29, Criteria1:=Array("确诊病例", "无症状感染者"), Operator:=xlFilterValues
    'Field:=29是目前最终分类
    ActiveWorkbook.Names.Add Name:="已确认阳性库!_FilterDatabase", RefersTo:="=已确认阳性库!$A:$AM", Visible:=False
    'Range("A1").Select
End Sub

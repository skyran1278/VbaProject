Sub Timesheet()

    '計時
    Time0 = Timer

    '關閉螢幕更新，加快速度
    'Application.ScreenUpdating = False

    Set Fs = CreateObject("Scripting.FileSystemObject")
    Dim Value(100000, 4) As String
    Dim HrType(3) As String

    Dim WbManage As Workbook
    Set WbManage = Workbooks("TimeSheetManageSystem")

    Dim WsEmployee As Worksheet
    Set WsEmployee = Worksheets("DateAndEmployee")

    Call DeletePriorData

    '激活DateAndEmployee
    Workbooks("TimeSheetManageSystem").Worksheets("DateAndEmployee").Activate

    '由於CurDir()會發生問題，所以由儲存格讀取路徑
    XlPath = Cells(6, 3)

    '陣列計數
    ValueRowNumber = -1

    ' 讀取資料成功計數
    ReadExcelSuccessRow = 6

    '讀取第4欄時間之列數
    TimeRowUsed = Cells(Rows.Count, 4).End(xlUp).Row

    '從第6列到最後一列
    For TimeRowNumber = 6 To TimeRowUsed

        '激活DateAndEmployee，避免出現Bug
        Workbooks("TimeSheetManageSystem").Worksheets("DateAndEmployee").Activate

        '讀取第5欄時間之列數
        EmployeeRowUsed = Cells(Rows.Count, 5).End(xlUp).Row

        '從第6列到最後一列
        For EmployeeRowNumber = 6 To EmployeeRowUsed

            '激活DateAndEmployee，避免出現Bug
            Workbooks("TimeSheetManageSystem").Worksheets("DateAndEmployee").Activate

            '讀取人名
            Employee = Cells(EmployeeRowNumber, 5)

            '結合Time和人名
            TimeAndEmployee = Cells(TimeRowNumber, 4) & " " & Employee

            '綜合以上得到檔案名稱
            DateAndEmployeePath = XlPath & "\" & Cells(TimeRowNumber, 4) & "\" & TimeAndEmployee & ".xlsx"

            '檢驗檔名是否存在 不存在的話就下一個人
            If Not (Fs.FileExists(DateAndEmployeePath)) Then GoTo NextEmployee

            ' 讀取資料成功紀錄
            Cells(ReadExcelSuccessRow, 11) = DateAndEmployeePath
            ReadExcelSuccessRow = ReadExcelSuccessRow + 1

            '如果存在檔案就打開他 並且不更新資料
            Workbooks.Open Filename:=DateAndEmployeePath, UpdateLinks:=0

            '第1週開始到第5週
            For WorksheetsNumber = 1 To 5

                '激活當週
                '有時候會出現陣列索引超出範圍，是因為Week5多一個空格，要刪掉多出的空格
                Workbooks(TimeAndEmployee).Worksheets("week" & WorksheetsNumber).Activate


                '讀取最後一列
                LastRowNumber = Cells(Rows.Count, 1).End(xlUp).Row

                '讀取資料Loop
                For CellsColumnNumber = 6 To 12
                    For CellsRowNumber = 9 To LastRowNumber - 1

                        '如果資料不為空且為數字 就記錄到陣列中
                        If Cells(CellsRowNumber, CellsColumnNumber) <> "" And IsNumeric(Cells(CellsRowNumber, CellsColumnNumber).Value) Then

                            '陣列計數
                            ValueRowNumber = ValueRowNumber + 1

                            'Employee
                            Value(ValueRowNumber, 0) = Employee

                            'Project
                            Value(ValueRowNumber, 1) = Cells(CellsRowNumber, 1)

                            'HrType
                            Value(ValueRowNumber, 2) = Cells(CellsRowNumber, 3)

                            'Worktime
                            Value(ValueRowNumber, 3) = Cells(CellsRowNumber, CellsColumnNumber)

                            'WorkDate
                            'DeBug，資料型態不符合，修正日期
                            Value(ValueRowNumber, 4) = Cells(6, CellsColumnNumber)
                        End If
                    Next
                Next
            Next

            '關掉檔案
            Workbooks(TimeAndEmployee).Close SaveChanges:=False

'不存在檔案，換下一個人
NextEmployee:
        Next
    Next

    '------------------------------------------------增加功能

    '工時總類
    HrType(0) = "Normal"
    HrType(1) = "Overtime"
    HrType(2) = "Overtime-H"

    '激活ProjectList
    WbManage.Worksheets("ProjectList").Activate

    '讀取第2欄時間之列數
    ProjectListRowUsed = Cells(Rows.Count, 2).End(xlUp).Row

    '從第2列到最後一列
    For ProjectRowNumber = 2 To ProjectListRowUsed

        '從第2列到最後一列
        For HrTypeNumber = 0 To 2

            '陣列記數
            ValueRowNumber = ValueRowNumber + 1

            'Project
            Value(ValueRowNumber, 1) = Cells(ProjectRowNumber, 2)

            'HrType
            Value(ValueRowNumber, 2) = HrType(HrTypeNumber)
        Next
    Next
    '------------------------------------------------

    Call DeleteDash(Value)

    '全部資料讀取完後
    '激活Database
    WbManage.Worksheets("Database").Activate

    '從第一筆讀最後一筆資料
    For PasteRowNumber = 0 To ValueRowNumber

        '由於是二維陣列，有兩個維度
        For PasteColumnNumber = 0 To 4

            '從第6行開始
            Cells(PasteRowNumber + 6, PasteColumnNumber + 3) = Value(PasteRowNumber, PasteColumnNumber)
        Next
    Next

    'Application.ScreenUpdating = True

    '關閉螢幕更新 測試耗時12秒13
    '開啟螢幕更新 測試耗時13秒79
    MsgBox "執行時間 " & Application.Round((Timer - Time0) / 60, 1) & " 分鐘", vbOKOnly

End Sub

Function DeleteDash(Value)

' 去掉 "-"

    For i = 1 To UBound(Value)
        Value(i, 1) = Replace(Value(i, 1), "-", " ")
    Next

End Function

Function DeletePriorData()

' 刪除之前的資料

    Workbooks("TimeSheetManageSystem").Worksheets("DateAndEmployee").Activate
    Columns(11).ClearContents
    Cells(5, 11) = "Success Read Excel"
    Workbooks("TimeSheetManageSystem").Worksheets("Database").Activate
    Range(Columns(3), Columns(7)).ClearContents
    Cells(5, 3) = "Employee"
    Cells(5, 4) = "Project"
    Cells(5, 5) = "HrType"
    Cells(5, 6) = "Worktime"
    Cells(5, 7) = "Workdate"

End Function

Sub Record()
'
' 巨集3 巨集
'

'
    Worksheets("ProjectRecord").Activate
    Worksheets("ProjectRecord").Cells.Delete
    Worksheets("Database").Activate
    ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
    SourceData:=Sheets("Database").Range(Cells(5, 3), Cells(5, 7)).CurrentRegion.Address).CreatePivotTable _
    TableDestination:=Worksheets("ProjectRecord").Cells(1, 1), TableName:="專案參與紀錄"
    Worksheets("ProjectRecord").Activate
    With ActiveSheet.PivotTables("專案參與紀錄")
        .PivotFields("Project").Orientation = xlRowField
        .PivotFields("HrType").Orientation = xlRowField
        .PivotFields("Employee").Orientation = xlColumnField
        .PivotFields("Worktime").Orientation = xlDataField
        .PivotFields("Project").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .TableStyle2 = "PivotStyleLight9"

    End With
End Sub







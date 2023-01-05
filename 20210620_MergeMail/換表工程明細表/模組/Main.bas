Attribute VB_Name = "Main"

Option Explicit
Dim ran As New Utils

Sub Main()
    ' input array and sort by 電號
    Dim inputSheet As Worksheet
    Set inputSheet = Worksheets("輸入")
    Dim inputArray As Variant
    inputArray = ran.GetRangeValues(inputSheet, 2, 1, ran.GetEndRowByCol(inputSheet, 3), 36)

    ' group by 計算日 and 營業區
    Dim outputCollection As New Collection
    Set outputCollection = GroupByElectricNumber(inputArray)

    ' clear output sheet
    Dim outputSheet As Worksheet
    Set outputSheet = Worksheets("輸出")
    outputSheet.Cells.Clear

    ' copy format range
    Dim srcSheet As Worksheet
    Set srcSheet = Worksheets("母版")
    srcSheet.Range("A1:BO41").Copy
    Dim i As Long
    For i = 0 To outputCollection.count - 1
        outputSheet.Range(outputSheet.Cells(1 + i * 41, 1), outputSheet.Cells(1 + (i + 1) * 41, 67)).PasteSpecial xlPasteAll
    Next i

    ' fill value
    Dim outputArray As Variant
    outputArray = ran.GetRangeValues(outputSheet, 1, 1, ran.GetEndRowByCol(outputSheet, 66), 67)
    outputArray = TaipowerModelToArray(outputCollection, outputArray)

    ' output array
    outputSheet.Range(outputSheet.Cells(1, 1), outputSheet.Cells(UBound(outputArray, 1), UBound(outputArray, 2))).value = outputArray

End Sub

Function TaipowerModelToArray(ByRef outputCollection As Collection, ByRef outputArray As Variant) As Variant
    Dim i As Long
    For i = 1 To outputCollection.count
        Dim taipowerModels As Variant
        taipowerModels = outputCollection.item(i)

        Dim row As Long
        row = (i - 1) * 41
        outputArray(row + 2, 46) = "頁數 " & i
        outputArray(row + 7, 2) = Mid(taipowerModels(1).CalculationDay, 1, 1)
        outputArray(row + 7, 3) = Mid(taipowerModels(1).CalculationDay, 2, 1)
        outputArray(row + 7, 4) = Mid(taipowerModels(1).District, 1, 1)
        outputArray(row + 7, 5) = Mid(taipowerModels(1).District, 2, 1)
        outputArray(row + 7, 6) = Mid(taipowerModels(1).BusinessArea, 1, 1)
        outputArray(row + 7, 7) = Mid(taipowerModels(1).BusinessArea, 2, 1)

        TaipowerModelToSubArray taipowerModels(1), row, outputArray
        TaipowerModelToSubArray taipowerModels(2), row + 11, outputArray
        TaipowerModelToSubArray taipowerModels(3), row + 22, outputArray

    Next i

    TaipowerModelToArray = outputArray
End Function

Sub TaipowerModelToSubArray(ByRef TaipowerModel As Variant, row As Long, ByRef outputArray As Variant)
' 將子表格的資料填入
    If IsEmpty(TaipowerModel) Then
        Exit Sub
    End If

    outputArray(row + 13, 1) = Mid(TaipowerModel.AccountNumber, 1, 1)
    outputArray(row + 13, 2) = Mid(TaipowerModel.AccountNumber, 2, 1)
    outputArray(row + 13, 3) = Mid(TaipowerModel.AccountNumber, 3, 1)
    outputArray(row + 13, 4) = Mid(TaipowerModel.AccountNumber, 4, 1)
    outputArray(row + 13, 5) = Mid(TaipowerModel.CategoryNumber, 1, 1)
    outputArray(row + 13, 6) = Mid(TaipowerModel.CategoryNumber, 2, 1)
    outputArray(row + 13, 7) = TaipowerModel.CheckNumber
    outputArray(row + 13, 17) = TaipowerModel.Matter
    outputArray(row + 14, 32) = TaipowerModel.UserName
    outputArray(row + 14, 57) = TaipowerModel.Coordinate & chr(10) & TaipowerModel.PoleNumber
    outputArray(row + 16, 32) = "用電地址: " & TaipowerModel.ElectricAddress
    outputArray(row + 17, 8) = Left(TaipowerModel.ElectricMeterNumber, 8)
    outputArray(row + 17, 19) = Mid(TaipowerModel.Type1, 1, 1)
    outputArray(row + 17, 20) = Mid(TaipowerModel.Type1, 2, 1)
    outputArray(row + 17, 21) = Mid(TaipowerModel.Ampere, 1, 1)
    outputArray(row + 17, 22) = Mid(TaipowerModel.Ampere, 2, 1)
    outputArray(row + 17, 23) = Mid(TaipowerModel.Multiple, 1, 1)
    outputArray(row + 17, 24) = Mid(TaipowerModel.Multiple, 2, 1)
    outputArray(row + 17, 25) = Mid(TaipowerModel.VerificationDeadline, 1, 1)
    outputArray(row + 17, 26) = Mid(TaipowerModel.VerificationDeadline, 2, 1)
    outputArray(row + 17, 27) = Mid(TaipowerModel.VerificationDeadline, 3, 1)
    outputArray(row + 17, 28) = Mid(TaipowerModel.VerificationDeadline, 5, 1)
    outputArray(row + 17, 29) = Mid(TaipowerModel.VerificationDeadline, 6, 1)
    outputArray(row + 17, 32) = "通訊地址: " & TaipowerModel.MailAddress
    outputArray(row + 18, 14) = Mid(TaipowerModel.CurrentValue, 1, 1)
    outputArray(row + 18, 15) = Mid(TaipowerModel.CurrentValue, 2, 1)
    outputArray(row + 18, 16) = Mid(TaipowerModel.CurrentValue, 3, 1)
    outputArray(row + 18, 17) = Mid(TaipowerModel.CurrentValue, 4, 1)
    outputArray(row + 18, 18) = Mid(TaipowerModel.CurrentValue, 5, 1)
    outputArray(row + 18, 19) = " (" & TaipowerModel.DifferentValue & ")"
    outputArray(row + 18, 32) = TaipowerModel.Phone1 & " " & TaipowerModel.Phone2
End Sub

Function GroupByElectricNumber(ByRef arr As Variant) As Collection
' 1. 按照電號排列
' 2. 連續且相同的計算日與營業區會變成同一頁
' 3. 但如果同一頁超過三個要變成下一頁
    Dim modelsCollection As New Collection

    ran.QuickSortArray arr, , , 4

    Dim previousKey As String
    previousKey = ""

    Dim subArray As Variant
    ReDim subArray(1 To 3)
    Dim subArrayIndex As Long
    subArrayIndex = 1

    Dim row As Long
    For row = LBound(arr, 1) To UBound(arr, 1)
        Dim newModel As TaipowerModel
        Set newModel = New TaipowerModel
        newModel.CalculationDay = arr(row, 3)
        newModel.ElectricNumber = arr(row, 4)
        newModel.Type1 = arr(row, 8)
        newModel.Matter = arr(row, 9)
        newModel.ElectricMeterNumber = arr(row, 10)
        newModel.Ampere = arr(row, 11)
        newModel.Multiple = arr(row, 12)
        newModel.VerificationDeadline = arr(row, 13)
        newModel.CurrentValue = arr(row, 17)
        newModel.NextDate = arr(row, 18)
        newModel.UserName = arr(row, 22)
        newModel.ElectricAddress = arr(row, 24)
        newModel.MailAddress = arr(row, 26)
        newModel.Phone1 = arr(row, 27)
        newModel.Phone2 = arr(row, 28)
        newModel.Coordinate = arr(row, 30)
        newModel.PoleNumber = arr(row, 31)
        newModel.DifferentValue = arr(row, 36)

        Dim newKey As String
        newKey = newModel.CalculationDay & "_" & newModel.BusinessArea
        If previousKey = newKey And subArrayIndex <= UBound(subArray) Then
            Set subArray(subArrayIndex) = newModel
        Else
            modelsCollection.Add subArray

            subArrayIndex = 1
            ReDim subArray(1 To 3)
            Set subArray(subArrayIndex) = newModel
        End If

        subArrayIndex = subArrayIndex + 1
        previousKey = newKey
    Next row

    ' 增加最後一筆資料並刪除第一筆空的資料
    modelsCollection.Remove 1
    modelsCollection.Add subArray

    Set GroupByElectricNumber = modelsCollection
End Function

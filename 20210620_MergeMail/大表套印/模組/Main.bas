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
    Set outputCollection = ArrayToTaipowerModel(inputArray)

    ' clear output sheet
    Dim outputSheet As Worksheet
    Set outputSheet = Worksheets("輸出")
    outputSheet.Cells.Clear
    'Delete all shapes in the active worksheet
    Dim shp As Shape
    For Each shp In outputSheet.Shapes
        shp.Delete
    Next shp

    ' copy format range
    Dim srcSheet As Worksheet
    Set srcSheet = Worksheets("母版")
    srcSheet.Rows("1:47").Copy
    Dim i As Long
    For i = 0 To outputCollection.count - 1
        outputSheet.Range(outputSheet.Cells(1 + i * 47, 1), outputSheet.Cells(1 + (i + 1) * 47, 95)).EntireRow.PasteSpecial Paste:=xlPasteFormats
    Next i

    srcSheet.Range("A1:CQ47").Copy
    For i = 0 To outputCollection.count - 1
        outputSheet.Paste Destination:=outputSheet.Cells(1 + i * 47, 1)
    Next i

    ' fill value
    Dim outputArray As Variant
    ' init size
    outputArray = ran.GetRangeValues(outputSheet, 1, 1, ran.GetEndRowByCol(outputSheet, 10), 95)
    outputArray = TaipowerModelToArray(outputCollection, outputArray)

    ' output array
    outputSheet.Range(outputSheet.Cells(1, 1), outputSheet.Cells(UBound(outputArray, 1), UBound(outputArray, 2))).value = outputArray

End Sub

Function TaipowerModelToArray(ByRef outputCollection As Collection, ByRef outputArray As Variant) As Variant
    Dim i As Long
    For i = 1 To outputCollection.count
        Dim taipowerModel As Variant
        Set taipowerModel = outputCollection.item(i)

        Dim row As Long
        row = (i - 1) * 47

        outputArray(row + 8, 66) = Mid(taipowerModel.CalculationDay, 1, 1)
        outputArray(row + 8, 67) = Mid(taipowerModel.CalculationDay, 2, 1)
        outputArray(row + 8, 68) = Mid(taipowerModel.District, 1, 1)
        outputArray(row + 8, 69) = Mid(taipowerModel.District, 2, 1)
        outputArray(row + 8, 70) = Mid(taipowerModel.BusinessArea, 1, 1)
        outputArray(row + 8, 71) = Mid(taipowerModel.BusinessArea, 2, 1)
        outputArray(row + 8, 72) = Mid(taipowerModel.AccountNumber, 1, 1)
        outputArray(row + 8, 73) = Mid(taipowerModel.AccountNumber, 2, 1)
        outputArray(row + 8, 74) = Mid(taipowerModel.AccountNumber, 3, 1)
        outputArray(row + 8, 75) = Mid(taipowerModel.AccountNumber, 4, 1)
        outputArray(row + 8, 76) = Mid(taipowerModel.CategoryNumber, 1, 1)
        outputArray(row + 8, 77) = Mid(taipowerModel.CategoryNumber, 2, 1)
        outputArray(row + 8, 78) = taipowerModel.CheckNumber
        outputArray(row + 8, 8) = taipowerModel.UserName
        outputArray(row + 12, 10) = taipowerModel.MailAddress
        outputArray(row + 12, 47) = taipowerModel.Phone1 & " " & taipowerModel.Phone2
        outputArray(row + 13, 79) = taipowerModel.Coordinate & chr(10) & taipowerModel.PoleNumber
        outputArray(row + 23, 16) = taipowerModel.Matter
        outputArray(row + 23, 17) = 0 ' Mid(taipowerModel.Ampere, 1, 1)
        outputArray(row + 23, 18) = 2 ' Mid(taipowerModel.Ampere, 2, 1)
        outputArray(row + 23, 19) = 0 ' Mid(taipowerModel.Ampere, 3, 1)
        outputArray(row + 23, 20) = 0 ' Mid(taipowerModel.Ampere, 4, 1)
        outputArray(row + 23, 33) = 4
        outputArray(row + 23, 34) = 0
        ' outputArray(row + 7, 8) = "用電地址: " & taipowerModel.ElectricAddress
        outputArray(row + 24, 14) = Mid(taipowerModel.Type1, 1, 1)
        outputArray(row + 24, 15) = Mid(taipowerModel.Type1, 2, 1)
        outputArray(row + 24, 17) = 0
        outputArray(row + 24, 18) = 2
        outputArray(row + 24, 19) = 0
        outputArray(row + 24, 20) = 0
        outputArray(row + 24, 22) = Mid(taipowerModel.ElectricMeterNumber, 1, 1)
        outputArray(row + 24, 23) = Mid(taipowerModel.ElectricMeterNumber, 2, 1)
        outputArray(row + 24, 24) = Mid(taipowerModel.ElectricMeterNumber, 3, 1)
        outputArray(row + 24, 25) = Mid(taipowerModel.ElectricMeterNumber, 4, 1)
        outputArray(row + 24, 26) = Mid(taipowerModel.ElectricMeterNumber, 5, 1)
        outputArray(row + 24, 27) = Mid(taipowerModel.ElectricMeterNumber, 6, 1)
        outputArray(row + 24, 28) = Mid(taipowerModel.ElectricMeterNumber, 7, 1)
        outputArray(row + 24, 29) = Mid(taipowerModel.ElectricMeterNumber, 8, 1)
        outputArray(row + 24, 33) = 4 ' Mid(taipowerModel.Multiple, 1, 1)
        outputArray(row + 24, 34) = 0 ' Mid(taipowerModel.Multiple, 2, 1)
        outputArray(row + 24, 36) = Mid(taipowerModel.VerificationDeadline, 1, 1)
        outputArray(row + 24, 37) = Mid(taipowerModel.VerificationDeadline, 2, 1)
        outputArray(row + 24, 38) = Mid(taipowerModel.VerificationDeadline, 3, 1)
        outputArray(row + 24, 39) = Mid(taipowerModel.VerificationDeadline, 5, 1)
        outputArray(row + 24, 40) = Mid(taipowerModel.VerificationDeadline, 6, 1)
        outputArray(row + 24, 41) = 0
        outputArray(row + 24, 42) = 0
        outputArray(row + 24, 43) = 0
        outputArray(row + 24, 72) = "W"

        outputArray(row + 30, 22) = Mid(taipowerModel.ElectricMeterNumber, 1, 1)
        outputArray(row + 30, 23) = Mid(taipowerModel.ElectricMeterNumber, 2, 1)
        outputArray(row + 30, 24) = Mid(taipowerModel.ElectricMeterNumber, 3, 1)
        outputArray(row + 30, 25) = Mid(taipowerModel.ElectricMeterNumber, 4, 1)
        outputArray(row + 30, 26) = Mid(taipowerModel.ElectricMeterNumber, 5, 1)
        outputArray(row + 30, 27) = Mid(taipowerModel.ElectricMeterNumber, 6, 1)
        outputArray(row + 30, 28) = Mid(taipowerModel.ElectricMeterNumber, 7, 1)
        outputArray(row + 30, 29) = Mid(taipowerModel.ElectricMeterNumber, 8, 1)
        outputArray(row + 30, 33) = 4 ' Mid(taipowerModel.Multiple, 1, 1)
        outputArray(row + 30, 34) = 0 ' Mid(taipowerModel.Multiple, 2, 1)
        outputArray(row + 30, 36) = Mid(taipowerModel.VerificationDeadline, 1, 1)
        outputArray(row + 30, 37) = Mid(taipowerModel.VerificationDeadline, 2, 1)
        outputArray(row + 30, 38) = Mid(taipowerModel.VerificationDeadline, 3, 1)
        outputArray(row + 30, 39) = Mid(taipowerModel.VerificationDeadline, 5, 1)
        outputArray(row + 30, 40) = Mid(taipowerModel.VerificationDeadline, 6, 1)
        ' outputArray(row + 18, 14) = Mid(taipowerModel.CurrentValue, 1, 1)
        ' outputArray(row + 18, 15) = Mid(taipowerModel.CurrentValue, 2, 1)
        ' outputArray(row + 18, 16) = Mid(taipowerModel.CurrentValue, 3, 1)
        ' outputArray(row + 18, 17) = Mid(taipowerModel.CurrentValue, 4, 1)
        ' outputArray(row + 18, 18) = Mid(taipowerModel.CurrentValue, 5, 1)
        ' outputArray(row + 18, 19) = " (" & taipowerModel.DifferentValue & ")"
        outputArray(row + 46, 95) = i

    Next i

    TaipowerModelToArray = outputArray
End Function

Function ArrayToTaipowerModel(ByRef arr As Variant) As Collection
    Dim modelsCollection As New Collection

    Dim row As Long
    For row = LBound(arr, 1) To UBound(arr, 1)
        Dim newModel As taipowerModel
        Set newModel = New taipowerModel
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

        modelsCollection.Add newModel
    Next row

    Set ArrayToTaipowerModel = modelsCollection
End Function

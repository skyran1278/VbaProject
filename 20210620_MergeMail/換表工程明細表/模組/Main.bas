Attribute VB_Name = "Main"

Option Explicit
Dim ran As New Utils

Sub Main()
    ' input array and sort by 電號
    Dim inputSheet As Worksheet
    Set inputSheet = Worksheets("輸入")
    Dim inputArray As Variant
    inputArray = ran.GetRangeValues(inputSheet, 2, 1, ran.GetEndRowByCol(inputSheet, 3), 36)
    ran.QuickSortArray inputArray, , , 4

    ' group by 計算日 and 營業區
    Dim dict As Object
    Set dict = GroupByElectricNumber(inputArray)

    ' The Collection to store the fixed-size arrays
    Dim outputCollection As New Collection
    Set outputCollection = GroupByFixedSizeChunk(dict, 3)

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
        With outputSheet.Range(outputSheet.Cells(1 + i * 41, 1), outputSheet.Cells(1 + (i + 1) * 41, 67))
            .PasteSpecial xlPasteAll
            ' .PasteSpecial xlPasteColumnWidths
            ' .PasteSpecial xlPasteValues, , False, False
            ' .PasteSpecial xlPasteFormats, , False, False
        End With
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
        outputArray(row + 7, 2) = Left(taipowerModels(1).CalculationDay, 1)
        outputArray(row + 7, 3) = Right(taipowerModels(1).CalculationDay, 1)
        outputArray(row + 7, 4) = Left(taipowerModels(1).District, 1)
        outputArray(row + 7, 5) = Right(taipowerModels(1).District, 1)
        outputArray(row + 7, 6) = Left(taipowerModels(1).BusinessArea, 1)
        outputArray(row + 7, 7) = Right(taipowerModels(1).BusinessArea, 1)

        TaipowerModelToArray2 taipowerModels(1), row, outputArray
        TaipowerModelToArray2 taipowerModels(2), row + 11, outputArray
        TaipowerModelToArray2 taipowerModels(3), row + 22, outputArray

    Next i

    TaipowerModelToArray = outputArray
End Function

Sub TaipowerModelToArray2(ByRef TaipowerModel As Variant, row As Long, ByRef outputArray As Variant)
    outputArray(row + 13, 1) = Left(TaipowerModel.AccountNumber, 1)
    outputArray(row + 13, 2) = Mid(TaipowerModel.AccountNumber, 2, 1)
    outputArray(row + 13, 3) = Mid(TaipowerModel.AccountNumber, 3, 1)
    outputArray(row + 13, 4) = Right(TaipowerModel.AccountNumber, 1)
    outputArray(row + 13, 5) = Left(TaipowerModel.CategoryNumber, 1)
    outputArray(row + 13, 6) = Right(TaipowerModel.CategoryNumber, 1)
    outputArray(row + 13, 7) = TaipowerModel.CheckNumber
    outputArray(row + 13, 17) = TaipowerModel.Matter
    outputArray(row + 14, 32) = TaipowerModel.UserName
    outputArray(row + 14, 57) = TaipowerModel.Coordinate & chr(10) & TaipowerModel.PoleNumber
    outputArray(row + 16, 32) = "用電地址: " & TaipowerModel.ElectricAddress
    outputArray(row + 17, 8) = Left(TaipowerModel.ElectricMeterNumber, 8)
    outputArray(row + 17, 19) = Left(TaipowerModel.Type1, 1)
    outputArray(row + 17, 20) = Right(TaipowerModel.Type1, 1)
    outputArray(row + 17, 21) = Left(TaipowerModel.Ampere, 1)
    outputArray(row + 17, 22) = Right(TaipowerModel.Ampere, 1)
    outputArray(row + 17, 23) = Left(TaipowerModel.Multiple, 1)
    outputArray(row + 17, 24) = Right(TaipowerModel.Multiple, 1)
    outputArray(row + 17, 25) = Left(TaipowerModel.VerificationDeadline, 1)
    outputArray(row + 17, 26) = Mid(TaipowerModel.VerificationDeadline, 2, 1)
    outputArray(row + 17, 27) = Mid(TaipowerModel.VerificationDeadline, 3, 1)
    outputArray(row + 17, 28) = Mid(TaipowerModel.VerificationDeadline, 5, 1)
    outputArray(row + 17, 29) = Mid(TaipowerModel.VerificationDeadline, 6, 1)
    outputArray(row + 17, 32) = "通訊地址: " & TaipowerModel.MailAddress
    outputArray(row + 18, 14) = Left(TaipowerModel.CurrentValue, 1)
    outputArray(row + 18, 15) = Mid(TaipowerModel.CurrentValue, 2, 1)
    outputArray(row + 18, 16) = Mid(TaipowerModel.CurrentValue, 3, 1)
    outputArray(row + 18, 17) = Mid(TaipowerModel.CurrentValue, 4, 1)
    outputArray(row + 18, 18) = Right(TaipowerModel.CurrentValue, 1)
    outputArray(row + 18, 19) = " (" & TaipowerModel.DifferentValue & ")"
    outputArray(row + 18, 32) = TaipowerModel.Phone1 & " " & TaipowerModel.Phone2
End Sub

Function GroupByFixedSizeChunk(ByRef dict As Object, chunkSize As Long) As Collection
    Dim outputCollection As New Collection

    Dim taipowerModels As Variant
    For Each taipowerModels In dict.Items
        ReDim Preserve taipowerModels(1 To ran.RoundUp(UBound(taipowerModels) / 3) * 3)
        Dim i As Long
        For i = LBound(taipowerModels) To UBound(taipowerModels) Step chunkSize
            ' Add the fixed-size array to the Collection
            Dim subArray As Variant
            ReDim subArray(1 To chunkSize)
            Dim j As Long
            For j = 1 To chunkSize
                Set subArray(j) = taipowerModels(j - 1 + i)
            Next j
            outputCollection.Add subArray
        Next
    Next taipowerModels

    Set GroupByFixedSizeChunk = outputCollection
End Function

Function GroupByElectricNumber(ByRef arr As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

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

        Dim key As String
        key = newModel.District & "_" & newModel.BusinessArea

        Dim taipowerModelCollection As Collection
        If dict.Exists(key) Then
            Set taipowerModelCollection = dict(key)
            taipowerModelCollection.Add newModel
        Else
            Set taipowerModelCollection = New Collection
            taipowerModelCollection.Add newModel
            dict.Add key, taipowerModelCollection
        End If
    Next row

    Dim key2 As Variant
    For Each key2 In dict.keys
        dict(key2) = collectionToArray(dict(key2))
    Next

    Set GroupByElectricNumber = dict
End Function

Function collectionToArray(collect As Collection) As Variant()
    Dim arr() As Variant
    ReDim arr(1 To collect.count)

    Dim i As Long
    For i = 1 To collect.count
        Set arr(i) = collect.item(i)
    Next

    collectionToArray = arr
End Function

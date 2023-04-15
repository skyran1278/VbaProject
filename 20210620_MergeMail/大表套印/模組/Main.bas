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
        Dim TaipowerModel As Variant
        Set TaipowerModel = outputCollection.item(i)

        Dim row As Long
        row = (i - 1) * 47


        outputArray(row + 8, 66) = Mid(TaipowerModel.CalculationDay, 1, 1)
        outputArray(row + 8, 67) = Mid(TaipowerModel.CalculationDay, 2, 1)
        outputArray(row + 8, 68) = Mid(TaipowerModel.District, 1, 1)
        outputArray(row + 8, 69) = Mid(TaipowerModel.District, 2, 1)
        outputArray(row + 8, 70) = Mid(TaipowerModel.BusinessArea, 1, 1)
        outputArray(row + 8, 71) = Mid(TaipowerModel.BusinessArea, 2, 1)
        outputArray(row + 8, 72) = Mid(TaipowerModel.AccountNumber, 1, 1)
        outputArray(row + 8, 73) = Mid(TaipowerModel.AccountNumber, 2, 1)
        outputArray(row + 8, 74) = Mid(TaipowerModel.AccountNumber, 3, 1)
        outputArray(row + 8, 75) = Mid(TaipowerModel.AccountNumber, 4, 1)
        outputArray(row + 8, 76) = Mid(TaipowerModel.CategoryNumber, 1, 1)
        outputArray(row + 8, 77) = Mid(TaipowerModel.CategoryNumber, 2, 1)
        outputArray(row + 8, 78) = TaipowerModel.CheckNumber
        outputArray(row + 8, 10) = TaipowerModel.UserName
        outputArray(row + 10, 10) = TaipowerModel.ElectricAddress
        outputArray(row + 12, 10) = TaipowerModel.MailAddress
        outputArray(row + 12, 47) = TaipowerModel.Phone1 & " " & TaipowerModel.Phone2
        outputArray(row + 13, 79) = TaipowerModel.Coordinate & chr(10) & TaipowerModel.PoleNumber
        outputArray(row + 23, 16) = TaipowerModel.Matter
        outputArray(row + 23, 17) = Mid(TaipowerModel.Ampere, 1, 1)
        outputArray(row + 23, 18) = Mid(TaipowerModel.Ampere, 2, 1)
        outputArray(row + 23, 19) = Mid(TaipowerModel.Ampere, 3, 1)
        outputArray(row + 23, 20) = Mid(TaipowerModel.Ampere, 4, 1)
        outputArray(row + 23, 33) = Mid(TaipowerModel.Multiple, 1, 1)
        outputArray(row + 23, 34) = Mid(TaipowerModel.Multiple, 2, 1)
        outputArray(row + 23, 72) = "W"



        outputArray(row + 24, 17) = 0
        outputArray(row + 24, 18) = 2
        outputArray(row + 24, 19) = 0
        outputArray(row + 24, 20) = 0
        outputArray(row + 24, 21) = 5
        outputArray(row + 24, 22) = Mid(TaipowerModel.ElectricMeterNumber, 1, 1)
        outputArray(row + 24, 23) = Mid(TaipowerModel.ElectricMeterNumber, 2, 1)
        outputArray(row + 24, 24) = Mid(TaipowerModel.ElectricMeterNumber, 3, 1)
        outputArray(row + 24, 25) = Mid(TaipowerModel.ElectricMeterNumber, 4, 1)
        outputArray(row + 24, 26) = Mid(TaipowerModel.ElectricMeterNumber, 5, 1)
        outputArray(row + 24, 27) = Mid(TaipowerModel.ElectricMeterNumber, 6, 1)
        outputArray(row + 24, 28) = Mid(TaipowerModel.ElectricMeterNumber, 7, 1)
        outputArray(row + 24, 29) = Mid(TaipowerModel.ElectricMeterNumber, 8, 1)
        outputArray(row + 24, 33) = Mid(TaipowerModel.Multiple, 1, 1)
        outputArray(row + 24, 34) = Mid(TaipowerModel.Multiple, 2, 1)
        outputArray(row + 24, 36) = Mid(TaipowerModel.VerificationDeadline, 1, 1)
        outputArray(row + 24, 37) = Mid(TaipowerModel.VerificationDeadline, 2, 1)
        outputArray(row + 24, 38) = Mid(TaipowerModel.VerificationDeadline, 3, 1)
        outputArray(row + 24, 39) = Mid(TaipowerModel.VerificationDeadline, 5, 1)
        outputArray(row + 24, 40) = Mid(TaipowerModel.VerificationDeadline, 6, 1)
        outputArray(row + 44, 9) = Mid(TaipowerModel.CurrentTransformer, 1, 1)
        outputArray(row + 44, 10) = Mid(TaipowerModel.CurrentTransformer, 2, 1)
        outputArray(row + 44, 11) = Mid(TaipowerModel.CurrentTransformer, 3, 1)
        outputArray(row + 44, 12) = Mid(TaipowerModel.CurrentTransformer, 4, 1)

        ' outputArray(row + 18, 19) = " (" & taipowerModel.DifferentValue & ")"
        outputArray(row + 46, 95) = i

        Dim tableTypes As Variant
        tableTypes = TaipowerModel.TableTypes

        Dim types As Variant
        types = TaipowerModel.Types
        Dim currentValues As Variant
        currentValues = TaipowerModel.CurrentValues
        Dim differentValues As Variant
        differentValues = TaipowerModel.DifferentValues

        Dim tableTypeIndex As Long
        For tableTypeIndex = LBound(tableTypes) To UBound(tableTypes)
            Dim tableTypeRow As Long
            Select Case tableTypes(tableTypeIndex)
                Case "01"
                    tableTypeRow = row + 23
                Case "02"
                    tableTypeRow = row + 25
                Case "03"
                    tableTypeRow = row + 27
                Case "04"
                    tableTypeRow = row + 29
                Case "06"
                    tableTypeRow = row + 31
                Case "08"
                    tableTypeRow = row + 33
                Case "09"
                    tableTypeRow = row + 35
                Case "10"
                    tableTypeRow = row + 37
                Case "11"
                    tableTypeRow = row + 39
                Case "12"
                    tableTypeRow = row + 41
                Case Else
                    MsgBox "出現未知的表別: " & tableTypes(tableTypeIndex)
            End Select
            outputArray(tableTypeRow, 8) = 1
            outputArray(tableTypeRow, 9) = 0
            outputArray(tableTypeRow, 10) = 1
            outputArray(tableTypeRow, 46) = currentValues(tableTypeIndex) & " (" & differentValues(tableTypeIndex) & ")"
            outputArray(tableTypeRow + 1, 8) = 1
            outputArray(tableTypeRow + 1, 9) = 0
            outputArray(tableTypeRow + 1, 10) = 1

            If tableTypeIndex <= UBound(types) Then
                outputArray(tableTypeRow + 1, 14) = Mid(types(tableTypeIndex), 1, 1)
                outputArray(tableTypeRow + 1, 15) = Mid(types(tableTypeIndex), 2, 1)
            End If

            outputArray(tableTypeRow + 1, 41) = 0
            outputArray(tableTypeRow + 1, 42) = 0
            outputArray(tableTypeRow + 1, 43) = 0

        Next tableTypeIndex
    Next i

    TaipowerModelToArray = outputArray
End Function

Function ArrayToTaipowerModel(ByRef arr As Variant) As Collection
    Dim modelsCollection As New Collection

    Dim row As Long
    For row = LBound(arr, 1) To UBound(arr, 1)
        Dim newModel As TaipowerModel
        Set newModel = New TaipowerModel
        newModel.CalculationDay = arr(row, 3)
        newModel.ElectricNumber = arr(row, 4)
        newModel.TableTypes = Split(arr(row, 7), " ")
        newModel.Types = Split(arr(row, 8), " ")
        newModel.Matter = arr(row, 9)
        newModel.ElectricMeterNumber = arr(row, 10)
        newModel.Multiple = arr(row, 12)
        newModel.VerificationDeadline = arr(row, 13)
        newModel.CurrentValues = Split(arr(row, 17), " ")
        newModel.NextDate = arr(row, 18)
        newModel.UserName = arr(row, 22)
        newModel.ElectricAddress = arr(row, 24)
        newModel.MailAddress = arr(row, 26)
        newModel.Phone1 = arr(row, 27)
        newModel.Phone2 = arr(row, 28)
        newModel.Coordinate = arr(row, 30)
        newModel.PoleNumber = arr(row, 31)
        newModel.DifferentValues = Split(arr(row, 36), " ")

        modelsCollection.Add newModel
    Next row

    Set ArrayToTaipowerModel = modelsCollection
End Function

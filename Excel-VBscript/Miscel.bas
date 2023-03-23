Attribute VB_Name = "Miscel"
Sub CreateBackup()
    'Create a backup copy of the current workbook
    Dim BackupFileName As String
    Dim WorkbookName As String
    Dim DateStamp As String
    
    WorkbookName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    DateStamp = Format(Now(), "yyyy-mm-dd_hhmmss")
    BackupFileName = ThisWorkbook.Path & "\" & WorkbookName & "_" & DateStamp & ".xlsm"
    
    ThisWorkbook.SaveCopyAs BackupFileName
End Sub

Sub PerformEverything()
    Call CreateBackup
    Call XpAbility.AddXPtoAll
    Call XpJob.AddJobXPtoAll
End Sub

Sub CancelEverything() ' minus deleting backup
    Call XpAbility.CancelAbilityXPtoAll
    Call XpJob.CancelJobXPtoAll
End Sub

Sub ChangeDateStatUpdate()
    
    ' Set value for todayValue
    Dim todayDate As Integer
    todayDate = ThisWorkbook.Worksheets("CharTable").Cells(6, 1)
    MsgBox "Debug: TodayDate is " + Str(todayDate)
    
    ' Dim otherWorkbook
    ' Dim otherSheet
    Dim otherWorkbook As Workbook
    Dim otherSheet As Worksheet
    
    ' Set otherWorkbook and otherWorksheet
    Set otherWorkbook = Workbooks.Open(ThisWorkbook.Path & "\ankleAttrbGrowth.xlsm")
    Set otherSheet = otherWorkbook.Worksheets("Notif")
    
    ' Define characterList
    Dim characterList As Variant
    'characterList = Array("Character0", "Character1")
    characterList = ExtractNonEmptyArray(1, 3, ThisWorkbook.Worksheets("CharTable"))
    
    ' Set otherWorkbook.OtherSheet.ToDay for each character
    Dim character As Variant
    For Each character In characterList
        ' Find character column
        Dim otherCharaCol As Integer
        otherCharaCol = otherWorkbook.Worksheets(otherSheet.Name).Evaluate("=CharacterColumn(""" & CStr(character) & """,""" & otherSheet.Name & """)")
        otherSheet.Cells(3, otherCharaCol) = todayDate - 1
        otherSheet.Cells(5, otherCharaCol) = todayDate
    Next character
    
    ' Update variable for each individual characters for each individual attributes
    character = "Character0" ' reset characterList counter
    Dim statList As Variant
    For Each character In characterList
        For i = 0 To 7
           Dim thisCharaCol As Integer
           thisCharaCol = CharacterColumn(CStr(character), "CharTable")
           otherCharaCol = otherWorkbook.Worksheets(otherSheet.Name).Evaluate("=CharacterColumn(""" & CStr(character) & """,""" & otherSheet.Name & """)")
           ThisWorkbook.Worksheets("Chartable").Cells(10 + i, thisCharaCol) = otherSheet.Cells(2 + i, otherCharaCol + 4)
       Next i
    Next character
    
    ' Close the other workbook
    otherWorkbook.Close
    
End Sub

Sub MoneyUpdate()
    Dim wsLog As Worksheet
    Dim wsDis As Worksheet
    Dim deltaSum As Currency
    
    ' Set references to the "Log" and "ChTbDis" sheets
    Set wsLog = ThisWorkbook.Worksheets("Log")
    Set wsDis = ThisWorkbook.Worksheets("ChTbDis")
    
    ' Loop through each row in column C of the "Log" sheets
    For i = 1 To wsLog.Range("C" & Rows.Count).End(xlUp).Row
        ' sum if column A is equal "delta"
        If wsLog.Range("A" & i).Value = "delta" Then
            deltaSum = deltaSum + wsLog.Range("C" & i).Value
        End If
    Next i
    
    ' add deltaSum to C32 in ChTbDis
    wsDis.Range("C32").Value = wsDis.Range("C32").Value + deltaSum
    
End Sub


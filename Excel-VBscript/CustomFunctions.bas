Attribute VB_Name = "CustomFunctions"
Function ConcatenateMultiple(Ref As Range, Optional Separator As String = " ") As String
Dim Cell As Range
Dim Result As String
For Each Cell In Ref
 Result = Result & Cell.Value & Separator
Next Cell
ConcatenateMultiple = Left(Result, Len(Result) - 1)
End Function

'Character is character name
Function AddJobXP(char As String)
    Dim characol As Integer
    Dim wss As Worksheet
    Dim wst As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim target As Range
    Dim AddedXP As Double
        
    Set wss = ThisWorkbook.Worksheets("Log")
    Set wst = ThisWorkbook.Worksheets("CharJobXP")
    
    'Find character column e.g "Character0" should return "2"
    characol = CharacterColumn(char, "CharJobXP")
    
    AddedXP = SumXP30x(char)
    
    ' Find the last row with data in column job name
    lastRow = wst.Cells(Rows.Count, characol).End(xlUp).Row
        
    ' Loop through each row
    For i = 2 To lastRow
        ' Check if the value in column note is "ACTIVE"
        Set target = wst.Cells(i, characol)
        If target.Offset(0, 6).Value = "ACTIVE" Then
            ' Add XP to the column "Experience"
            'target.Offset(0, 2).Value = target.Offset(0, 2).Value + wss.Range("Q8").Value
            
            target.Offset(0, 2).Value = target.Offset(0, 2).Value + AddedXP
        End If
    Next i
End Function

Function SubstractJobXP(char As String)
    Dim characol As Integer
    Dim wss As Worksheet
    Dim wst As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim target As Range
        
    Set wss = ThisWorkbook.Worksheets("Log")
    Set wst = ThisWorkbook.Worksheets("CharJobXP")
    
    'Find character column e.g "Character0" should return "B"
    characol = CharacterColumn(char, "CharJobXP")
    
    'Add "delta" Exp in .log multiplied by 30
    AddedXP = SumXP30x(char)
    
    ' Find the last row with data in column job name
    lastRow = wst.Cells(Rows.Count, characol).End(xlUp).Row
    
    ' Loop through each row
    For i = 2 To lastRow
        ' Check if the value in column note is "ACTIVE"
        Set target = wst.Cells(i, characol)
        If target.Offset(0, 6).Value = "ACTIVE" Then
            ' Add XP to the column "Experience"
            'target.Offset(0, 2).Value = target.Offset(0, 2).Value - wss.Range("Q8").Value
            
            target.Offset(0, 2).Value = target.Offset(0, 2).Value - AddedXP
        End If
    Next i
End Function

Function CharacterColumn(charaname As String, wss As String) As Integer
    Dim searchValue As String
    Dim searchRange As Range
    Dim foundColumn As Range
    Dim lastColumn As Long
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(wss)
    
    'set search value
    searchValue = charaname
    
    ' Find the last column with data in the topmost row
    lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Set search range to the range of nonempty columns in row 1
    Set searchRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastColumn))
        
    ' Use find method to find the first cell in the range that contains searchValue
    Set foundColumn = searchRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' check if the value is in the range
    If Not foundColumn Is Nothing Then
        ' Get Column number of found Cell and return it
        CharacterColumn = foundColumn.Column
    Else
        ' display if not found
        MsgBox "Character " + charaname + " is not found"
    End If
      
    
End Function

Function SumXP30x(charname As String) As Double
    Dim wss As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim total As Double
    Dim characol As Integer
    
    Set wss = ThisWorkbook.Worksheets("Log")
    
    ' search character column i.e "Character0", "Character1", etc in wss
    With wss
   
        characol = CharacterColumn(charname, "Log")
                
    End With
    
    ' search last row of character TotXP (offset +3 from character column)
    lastRow = wss.Cells(Rows.Count, characol + 3).End(xlUp).Row
    
    For i = 2 To lastRow
        
        If wss.Cells(i, 1).Value = "delta" Then ' If "delta"
            
            total = total + wss.Cells(i, characol + 2).Value ' sum over value in column "Exp"
            
        End If
                          
    Next i
    
    SumXP30x = 30 * total
    
    MsgBox "Net addedXP to be added to character " + charname + " is " + Str(SumXP30x)
    
End Function


Function AddAbilityXP(char As String)
    Dim characol As Integer
    Dim wss As Worksheet
    Dim wst As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim target As Range
    Dim AddedXP As Double
        
    Set wss = ThisWorkbook.Worksheets("Log")
    Set wst = ThisWorkbook.Worksheets("CharAbilityXP")
    
    'Find character column e.g "Character0" should return "2"
    characol = CharacterColumn(char, "CharAbilityXP")
    
    AddedXP = SumXP30x(char)
    
    ' Find the last row with data in column job name
    lastRow = wst.Cells(Rows.Count, characol).End(xlUp).Row
        
    ' Loop through each row
    For i = 2 To lastRow
        ' Check if the value in column note is "delta"
        Set target = wst.Cells(i, characol)
        ' check note for delta (offset +5 from charactercolumn)
        If target.Offset(0, 5).Value = "delta" Then
            ' Add XP to the column "Experience" (offset column + 2)
            'target.Offset(0, 2).Value = target.Offset(0, 2).Value + wss.Range("Q8").Value
            
            target.Offset(0, 2).Value = target.Offset(0, 2).Value + AddedXP
        End If
    Next i
End Function

Function SubstractAbilityXP(char As String)
    Dim characol As Integer
    Dim wss As Worksheet
    Dim wst As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim target As Range
    Dim AddedXP As Double
        
    Set wss = ThisWorkbook.Worksheets("Log")
    Set wst = ThisWorkbook.Worksheets("CharAbilityXP")
    
    'Find character column e.g "Character0" should return "2"
    characol = CharacterColumn(char, "CharAbilityXP")
    
    AddedXP = SumXP30x(char)
    
    ' Find the last row with data in column job name
    lastRow = wst.Cells(Rows.Count, characol).End(xlUp).Row
        
    ' Loop through each row
    For i = 2 To lastRow
        ' Check if the value in column note is "delta"
        Set target = wst.Cells(i, characol)
        ' check note for delta (offset +5 from charactercolumn)
        If target.Offset(0, 5).Value = "delta" Then
            ' Add XP to the column "Experience" (offset column + 2)
            'target.Offset(0, 2).Value = target.Offset(0, 2).Value + wss.Range("Q8").Value
            
            target.Offset(0, 2).Value = target.Offset(0, 2).Value - AddedXP
        End If
    Next i
End Function

Function ExtractNonEmptyArray(rowNum As Long, firstCol As Long, ws As Worksheet) As Variant
    Dim myArray() As Variant ' Declare Array
    Dim lastCol As Long ' Declare last nonempty column
    Dim i As Long ' Loop variable
    
    'set the range from the firstCol to the last used column in the row
    lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    Set myRange = ws.Range(ws.Cells(rowNum, firstCol), ws.Cells(rowNum, lastCol))
    
    'loop through each columns in the row and add the nonempty value to the array
    For Each Cell In myRange
        If Not IsEmpty(Cell) Then
            ReDim Preserve myArray(i) ' resize array to accomodate new value
            myArray(i) = Cell.Value
            i = i + 1
        End If
    Next Cell
    
    ' Return the array
    ExtractNonEmptyArray = myArray
    
End Function

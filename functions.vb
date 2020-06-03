Public Function ConcatIf(criteriaRange As Range, ifCriteria As Variant, Optional concatRange As Range, Optional concatSpacer As String = "") As String
'Concatenates cells in concatRange based on whether their sister-cell in criteriaRange meets the ifCriteria
'Follows the same rules as SUMIF as best as possible
'Works on Multiple dimension ranges
'Will use the criteriaRange as the concatRange if no concatRange is specified
'concatRange is basically only relevant for its top-left cell index:
'    if we are looking at a criteriaRange cell at 2,3 then we will
'    concat the cell that is 2,3 relative to the concatRange top-left cell
    
    Dim rowStart As Integer
    Dim colStart As Integer
    rowStart = criteriaRange.Row
    colStart = criteriaRange.Column
    
    Dim evOp As Variant
    evOp = Left$(ifCriteria, 1)
    
    Dim result As String
    Dim val As Variant
    Dim matched As String
    Dim rowRel As Integer
    Dim colRel As Integer
    
    'If the user did not define concatRange
    If concatRange Is Nothing Then
        'Default it to the criteriaRange
        Set concatRange = criteriaRange
    End If
    
    For Each entry In criteriaRange
        'Get the criteria evaluation string
        val = entry.Value
        If IsEmpty(val) Then
            val = """"""
        End If
        evalString = """" & val & """=""" & ifCriteria & """"

        matched = False
        'If the the criteriaRange cell matches the criteria
        If evOp = "=" Or evOp = "<" Or evOp = ">" Then
            If Not IsNumeric(val) Then
                val = """" & val & """"
            End If
            If Application.Evaluate(val & ifCriteria) Then
                matched = True
            End If
        Else
            If val Like ifCriteria Then
                matched = True
            End If
        End If

        If matched Then
            rowRel = entry.Row - rowStart + 1
            colRel = entry.Column - colStart + 1
            result = result & concatRange.Cells(rowRel, colRel).Value & concatSpacer
        End If
    Next entry
    
    ConcatIf = result
End Function

Private Sub test_concatif()
    Debug.Print "---"
    Debug.Print "RESULT:  " & ConcatIf(Range("H12:H18"), 1, Range("C12:L18"))
    Debug.Print "RESULT:  " & ConcatIf(Range("H12:H18"), "=11", Range("C12:L18"))
    Debug.Print "RESULT:  " & ConcatIf(Range("H12:H18"), ">=12", Range("C12:L18"))
    Debug.Print "RESULT:  " & ConcatIf(Range("H12:H18"), ">12", Range("C12:L18"))
    Debug.Print "RESULT:  " & ConcatIf(Range("H12:H18"), "p?p", Range("C12:L18"))
    Debug.Print "RESULT:  " & ConcatIf(Range("H12:H18"), "p*p", Range("C12:L18"))
    Debug.Print "RESULT:  " & ConcatIf(Range("H12:H18"), "<>1", Range("C12:L18"))
End Sub

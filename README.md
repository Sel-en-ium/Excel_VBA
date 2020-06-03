## Developing
* Excel > Developer Tab > Visual Basic
    * A new "Microsoft Visual Basic" window should open
* Right click "ThisWorkbook" (on the left side project explorer), Insert > New Module
* Define stuff in the module!

### Debugging

Trigger code:  
* Excel > Developer Tab > Insert > Command Button (Activex Control)
    * Place the button somewhere
* Excel > Developer Tab (ensure you are in "Design Mode")
* Right click on button, "View code"
    * Here you can call your function on each button press with some default values

```VBA
Private Sub CommandButton1_Click()
    Dim res As String
    Debug.Print concatIf(Range("N15:O18"), 1)
    Debug.Print ConcatIf(Range("N15:O18"), 1, Range("K15:L18"))
End Sub
```

Immediate window: 
* Press ctrl + G
* you can type interactive commands in here
    * Note: Often you need to preface the command with `Debug.Print` or you will get an error


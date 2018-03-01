# VBA
My VBA Collect Code

# VBA Cheet Sheet

## File Process

1. Loop through files in a folder
```VBA
Sub LoopThroughFiles()
    Dim fileName As String
    fileName = Dir("C:\SearchFolder\*patternSearch*")
    Do While Len(fileName) > 0
        Debug.Print fileName
        fileName = Dir
    Loop
End Sub
```

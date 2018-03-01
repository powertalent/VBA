
# VBA Cheet Sheet

## File Process

1. Loop through files in a folder
```VBA
Sub LoopThroughFiles()
    Dim StrFile As String
    StrFile = Dir("c:\testfolder\*test*")
    Do While Len(StrFile) > 0
        Debug.Print StrFile
        StrFile = Dir
    Loop
End Sub
```

2. Read Data From Close File
```VBA
Sub ReadDataFromCloseFile()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim src As Workbook
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open("C:\Q-SALES.xlsx", True, True)
    
    ' PROCESSING....
    
    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
```

## Range Process

1. For Each cell in a range
```VBA
For
```

## Shape Process

1. For Each Shape in a ActiveSheet
```VBA
For Each shape In ActiveSheet.Shapes
    ' PROCESSING...
Next shape
```

## SORT

1. QuickSort
```VBA
  Sub QuickSort(arr, Lo As Long, Hi As Long)
  	Dim varPivot As Variant
    Dim varTmp As Variant
    Dim tmpLow As Long
    Dim tmpHi As Long
    tmpLow = Lo
    tmpHi = Hi
    varPivot = arr((Lo + Hi) \ 2)
    Do While tmpLow <= tmpHi
      Do While arr(tmpLow) < varPivot And tmpLow < Hi
        tmpLow = tmpLow + 1
      Loop
      Do While varPivot < arr(tmpHi) And tmpHi > Lo
        tmpHi = tmpHi - 1
      Loop
      If tmpLow <= tmpHi Then
        varTmp = arr(tmpLow)
        arr(tmpLow) = arr(tmpHi)
        arr(tmpHi) = varTmp
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
      End If
    Loop
    If Lo < tmpHi Then QuickSort arr, Lo, tmpHi
    If tmpLow < Hi Then QuickSort arr, tmpLow, Hi
  End Sub
```


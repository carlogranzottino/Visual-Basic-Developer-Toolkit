Attribute VB_Name = "XDW_Toolkit"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''      XDW_Toolkit
''
''          ~For frequently used custom functions and procedures
''
''          ~Created by David Wilson | Any code adapted from other sources is credited in the respective procedure
''
''          ~Last updated: 07/25/2019
''
''      --------------------------------------------------------------------------''
''
''       Content Procedures and Functions:
''
''          StartWrapper
''          EndWrapper
''          FindLastRow
''          FindFullRange
''          ImportFileData
''          IsFileOpen
''          PrintArray
''          IsArrayOneDimensional
''          NumberOfArrayDimensions
''          SetFilePath
''          TrapTrim
''          ReturnColumnLetter
''          ReturnColumnNumber
''          CreateSingleTabWorkbook
''          ShowLabelNames
''          LogTime
''          ColorIndexToRGB
''          SortData
''          Levenshtein
''          ShowLabelNames
''          ReadTextFile
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): StartWrapper; EndWrapper
'used for: bookending modules for performance optimization purposes
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StartWrapper()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub


Sub EndWrapper()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = False
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): FindLastRow
'used for: finding the last row in a given column
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function FindLastRow(ThisWorksheet As Worksheet, ThisColumn As String) As Double
    
    With ThisWorksheet
        FindLastRow = .Range(ThisColumn & .Rows.count).End(xlUp).Row
    End With

End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): FindFullRange
'used for: finding the full range of a given sheet; can find last row or column; variants (arrays) can be assigned to the function output
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function FindFullRange(ThisWorksheet As Worksheet, Optional ByRef LastRow As Double, Optional ByRef LastColumn As Double) As Range

    Dim usedRange As Range
    '''''
    With ThisWorksheet
        If ThisWorksheet.Application.CountA(Cells) = 0 Then 'if selected sheet is blank
            LastRow = 1
            LastColumn = 1
        Else
            LastRow = .Cells.Find(What:="*", after:=.Range("A1"), LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
            LastColumn = .Cells.Find(What:="*", after:=.Range("A1"), LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
        End If
        '''''
        Set FindFullRange = .Range(.Cells(1, 1), .Cells(lastRow, LastColumn))
    End With

End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): ImportFileData; IsFileOpen
'used for: Importing Excel file data from a worksheet in another workbook; it is returned at a two-dimensional array.  
'calls other XDW tools: yes - ImportFileData requires IsFileOpen for validation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ImportFileData(filepath As String, sheetName As String, keepImportFileOpen As Boolean) As Variant
    
    Dim wb As Workbook
    Dim alreadyOpen As Boolean
    ''''''
    If Len(Dir(filepath)) < 1 Then
        MsgBox "The file expected at: '" & filepath & "' could not be found.  Please ensure that the file is saved to this location and try again."
        EndWrapper
        End
    End If
    ''''''
    DoEvents
    ''''''
    alreadyOpen = IsFileOpen(filepath)
    ''''''
    Set wb = Workbooks.Open(filepath, ReadOnly:=True)
    ImportFileData = FindFullRange(wb.Sheets(sheetName))
    If keepImportFileOpen = False And alreadyOpen = True Then wb.Close
    
End Function

Function IsFileOpen(FileName As String) As Boolean

    Dim fileBum As Integer
    Dim errNum As Integer
    ''''''
    On Error Resume Next                            ' Turn error checking off.
    fileNum = FreeFile()                            ' Get a free file number.
    Open FileName For Input Lock Read As #fileNum
    Close fileNum                                   ' Close the file.
    errNum = Err                                    ' Save the error number that occurred.
    On Error GoTo 0                                 ' Turn error checking back on.
    ''''''
    Select Case errNum
        Case 0  ' No error occurred; File is not already open by another user
            IsFileOpen = False
        Case 70 ' Error number for "Permission Denied."; File is already opened by another user.
            IsFileOpen = True
        Case Else
            Error errNum
    End Select
    
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): PrintArray; IsArrayOneDimensional; NumberOfArrayDimensions
'used for: printing a given variant to a designated location on worksheet
'calls other XDW tools: yes - PrintArry requires the other two subs for data checks to prevent errors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PrintArray(printCell As Range, dataArray As Variant)

    Dim aHeight As Double
    Dim aWidth As Double
    Dim rngPrint As Range
    Dim printData As Variant
    Dim adjZeroX As Integer
    Dim adjZeroY As Integer
    '''''
    printData = dataArray
    '''''
    If IsEmpty(printData) = True Then ReDim printData(0)
    '''''
    If IsArrayOneDimensional(printData) = True Then
        aHeight = UBound(printData)
        aWidth = 1
    Else
        If LBound(dataArray, 1) = 0 Then adjZeroX = 1
        If LBound(dataArray, 2) = 0 Then adjZeroY = 1
        aHeight = UBound(printData, 1) + adjZeroX
        aWidth = UBound(printData, 2) + adjZeroY
    End If
    '''''
    Set rngPrint = printCell
    Set rngPrint = rngPrint.Resize(aHeight, aWidth)
    rngPrint = printData

End Sub

Function IsArrayOneDimensional(arr As Variant) As Boolean
    'This function adapted from code by Chip Pearson
    'URL: http://www.cpearson.com/excel/vbaarrays.htm
    
    IsArrayOneDimensional = (NumberOfArrayDimensions(arr) = 1)
    
End Function

Function NumberOfArrayDimensions(thisArr As Variant) As Integer
    'This function created by Chip Pearson
    'URL: http://www.cpearson.com/excel/vbaarrays.htm
    
    Dim dimensionCount As Integer
    Dim currDim As Integer
    '''''
    On Error Resume Next
    '''''
    Do
        dimensionCount = dimensionCount + 1
        currDim = UBound(thisArr, dimensionCount)
    Loop Until Err.Number <> 0
    ''''
    NumberOfArrayDimensions = dimensionCount - 1
    
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): SetFilePath
'used for: creates file dialog in Windows Explorer that allows the user to pick the new file path and then sets the path to the function as a string
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function SetFilePath(DisplayName As String, BoxTitle As String, oldPath As String) As String
    
    Dim fDialog As Office.FileDialog

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    With fDialog
        .InitialFileName = oldPath
        .AllowMultiSelect = False
        .Title = BoxTitle
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        If .Show = True Then
            SetFilePath = .SelectedItems(1)
        Else
            SetFilePath = oldPath
        End If
   End With
   
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): TrapTrim
'used for: trimming imported data that may contain error values or data mismatches -- typically this is from sheets used by a number of users or external systems
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function TrapTrim(ByVal inputValue As Variant, errorValue As String, Optional makeLowerCase As Boolean) As String

    On Error GoTo ErrorTrap
    '''''
    If IsError(inputValue) = True Then
        TrapTrim = errorValue
    ElseIf IsNull(inputValue) = True Then
        TrapTrim = errorValue
    Else
        TrapTrim = Trim(inputValue)
        If makeLowerCase = True Then TrapTrim = LCase(TrapTrim)
    End If
    '''''
    Exit Function
        
ErrorTrap:
    TrapTrim = errorValue

End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): ReturnColumnLetter; ReturnColumnNumber
'used for: converting a column identifier to the other type
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ReturnColumnLetter(ByVal columnNumber As Double) As String

    ReturnColumnLetter = Split(Cells(, columnNumber).Address, "$")(1)
    
End Function

Function ReturnColumnNumber(ByVal columnLetter As String) As Double

    ReturnColumnNumber = Range(columnLetter & 1).Column

End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): CreateSingleSheetWorkbook
'used for: to create a new workbook containing just one worksheet  -- the build in functionality for Excel "SheetsInNewWorkbook" has caused errors with certain versions
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CreateSingleTabWorkbook(sheetName As String) As Workbook
    'VBA currently has native functionality that does this using 'SheetsInNewWorkbook'
    'This isn't supported by older versions of VBA
    
    Dim SelectedSheet As Worksheet

    Dim wb As Workbook
    
    Set wb = Workbooks.Add
    sheetName = Trim(sheetName)
    If sheetName = "" Then sheetName = "Sheet1"
    wb.Sheets(1).Name = sheetName
    
    For Each SelectedSheet In wb.Sheets
        If Not SelectedSheet.Name = sheetName Then SelectedSheet.Delete
    Next SelectedSheet
    
    Set CreateSingleTabWorkbook = wb

End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): SheetExists
'used for: verifying if worksheet exists in a given workbook
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sh As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sh = wb.Sheets(shName)
     On Error GoTo 0
     SheetExists = Not sh Is Nothing
 End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): LogTime
'used for: create time log for testing the run time duration of procedures
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function LogTime(ByRef startTime As Double, TaskName As String) As Double

    Dim RunTime As Double
    
    RunTime = Timer - startTime
    Debug.Print TaskName & " Run Time:  " & RunTime & " seconds"
    LogTime = Timer
    
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): ColorIndexToRGB
'used for: Convert a given cell's color index (found via property ".color.interior") to RBG to be applied to other Excel objects
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ColorIndexToRGB(ByRef ColorIndex As Variant, ByRef R As Variant, ByRef G As Variant, ByRef B As Variant)
    R = ColorIndex Mod 256
    G = ColorIndex \ 256 Mod 256
    B = ColorIndex \ 65536 Mod 256
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): SortSheet
'used for: Sorting a worksheet table
'calls other XDW tools: yes - FindFullRange() is used to find last row and column if not specified for user; ReturnColumnLetter() is used to convert the column number to letter
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SortData(SortSheet As Worksheet, SortColumnLetter As String, SortMethod As String, Optional HasHeader As Boolean = True, Optional ByVal lastRow As Double = -9999, Optional ByVal LastColumn As Double = -9999)
    'SORT METHODS CAN BE EITHER "asc","ascending","desc", or "descending"; they are not case sensitive
        
        Dim startRow As Double
        Dim ynHeader As Variant
        Dim sortOrder As Variant
        
        If (lastRow = -9999) And (LastColumn = -9999) Then
            FindFullRange SortSheet, lastRow, LastColumn
        ElseIf (lastRow = -9999) Then
            FindFullRange SortSheet, lastRow
        ElseIf (LastColumn = -9999) Then
            FindFullRange SortSheet, , LastColumn
        End If
        
        If HasHeader = True Then
            startRow = 2
            ynHeader = xlYes
        Else
            startRow = 1
            ynHeader = xlNo
        End If
        
        Select Case UCase(SortMethod)
            Case "ASC", "ASCENDING"
                sortOrder = xlAscending
            Case "DESC", "DESCENDING"
                sortOrder = xlDescending
            Case Else
                Debug.Print "NO CORRECT SORT TYPE COULD BE FOUND. ERROR WILL BE THROWN."
        End Select
        
        With SortSheet.Sort
                .SortFields.Clear
                .SortFields.Add key:=SortSheet.Range(SortColumnLetter & startRow & ":" & SortColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=sortOrder, DataOption:=xlSortNormal
                .SetRange SortSheet.Range("A1:" & ReturnColumnLetter(LastColumn) & lastRow)
                .Header = ynHeader
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
        End With
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): Levenshtein
'used for: finding the Levenshtein distance between two strings
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Levenshtein(s1 As String, s2 As String) As Integer
    'this code was adapted from a post on StackOverflow from user "smirkingman"
    'URL: https://stackoverflow.com/questions/4243036/levenshtein-distance-in-vba/6423088
    
    Dim i As Integer, _
        j As Integer, _
        len1 As Integer, _
        len2 As Integer, _
        d() As Integer, _
        min1 As Integer, _
        min2 As Integer
    
    len1 = Len(s1)
    len2 = Len(s2)
    ReDim d(len1, len2)
    For i = 0 To len1
        d(i, 0) = i
    Next
    For j = 0 To len2
        d(0, j) = j
    Next
    For i = 1 To len1
        For j = 1 To len2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                d(i, j) = d(i - 1, j - 1)
            Else
                min1 = d(i - 1, j) + 1
                min2 = d(i, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                min2 = d(i - 1, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                d(i, j) = min1
            End If
        Next
    Next
    Levenshtein = d(len1, len2)
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): ShowLabelNames
'used for: procedure used toggle a given workbooks label names between not visible and visible; used to hide named ranges from non-dev users
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowLabelNames(ShowNames As Boolean)

    Dim cellName As Name
    
    For Each cellName In ThisWorkbook.Names
        If ShowNames = False Then
            If cellName.Visible = True Then cellName.Visible = False
        Else
            If cellName.Visible = False Then cellName.Visible = True
        End If
    Next cellName

End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): ReadTextFile
'used for: reading the contexts of a text file into an array
'calls other XDW tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function ReadTextFile(filepath As String, DELIMITER As String) As Variant
    'Credit: this code was adapted from a page on TheSpreadsheetGuru
    'URL: https://www.thespreadsheetguru.com/blog/vba-guide-text-files

    Dim textFile As Integer
    
    Dim fileContent As String
    Dim lineArr() As String
    Dim wordArr() As String
    Dim dataArray() As Variant
    
    Dim x As Long
    Dim y As Long
    
         
    'Open the text file in a Read State
    textFile = FreeFile
    Open filepath For Input As textFile
    
    fileContent = Input(LOF(textFile), textFile)
    Close textFile
      
    'Separate Out lines of data
    lineArr = Split(fileContent, vbCrLf)
    wordArr = Split(lineArr(LBound(lineArr)), vbTab)
    ReDim dataArray(UBound(lineArr), UBound(wordArr))
    
    'Read Data into an Array Variable
    For x = LBound(lineArr) To UBound(lineArr)
        If Len(Trim(lineArr(x))) <> 0 Then
            'Split up line of text by delimiter
            wordArr = Split(lineArr(x), vbTab)
            
            'Load line of data into Array variable
            For y = LBound(wordArr) To UBound(wordArr)
              dataArray(x, y) = Trim(wordArr(y))
            Next y
        End If
    Next x

    ReadTextFile = dataArray

End Function








'-----------------------End Toolbox-----------------------


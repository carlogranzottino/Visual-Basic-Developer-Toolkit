Attribute VB_Name = "XDW_DictionaryTools"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''      XDW_Toolkit
''
''          ~For custom dictionary functions
''
''          ~Created by David Wilson
''
''          ~Last updated: 01/30/2018
''
''      --------------------------------------------------------------------------''
''
''       Content Procedures and Functions:
''
''          CreateHeaderReference
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): CreateHeaderReference
'used for: create reference dictionary to find column header locations for files in which the order changes sometimes
'calls other tools: yes: [XDW_Toolkit=>] TrapTrim; FindFullRange
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CreateHeaderReference(RefSheet As Worksheet, HeaderRow As Long) As Dictionary
    'PLEASE NOTE: this requires that the reference to "Microsoft Scripting Runtime" is checked (Tools-->References)
    Dim refDict As Dictionary
    Dim refData As Variant
    Dim xHead As String
    Dim y As Double
    
    Set refDict = New Dictionary
    refData = FindFullRange(RefSheet)
    
    For y = LBound(refData, 2) To UBound(refData, 2)
            xHead = TrapTrim(tData(HeaderRow, y), "", True)
            refDict.Item(xHead) = y
    Next y
    
    Set CreateHeaderReference = refDict
    
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure(s): PrintDictionary
'used for: Prints dictionary to a specified column.  The variable 'printCell' must be a one cell range.  It is optional to print the values associated with each key.
'calls other tools: no
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PrintDictionary(printCell As Range, printDict As Dictionary, Optional printKeyVal As Boolean = False)
    
    
    Dim key As Variant
    Dim rngK As Range, _
        rngV As Range

    Set rngK = printCell
    Set rngV = printCell.Offset(0, 1)
    

    For Each key In printDict
        rngK.Value = key
        If printKeyVal Then
            rngV.Value = printDict(key)
        End If
        Set rngK = rngK.Offset(1, 0)
        Set rngV = rngV.Offset(1, 0)
    Next key

End Sub



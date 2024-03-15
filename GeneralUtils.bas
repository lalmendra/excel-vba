Attribute VB_Name = "GeneralUtils"
Option Explicit
'####################################################################

'## Naming conventions

'# Prefixes:

'ai = Array Index
'ali = Array Last Index
'arr = Array
'c = Column
'ch = Column Header
'dt = Data type
'eh = Error Handler
'f = Formatting
'fc = First Column
'fr = First Row
'hr = Header Row (Row in a worksheet where the Header[Column name] is)
'i = Iterator / Index
'lc = Last Column
'lr = Last Row
'mc = Main Column (A column that will have data in every record, normally its the ID column)
'r = Row
'wb = Workbook
'ws = Worksheet
'nws = Worsheet Name

'####################################################################

'## Standard Public constants

'# Standard Header Row and Main Column
Public Const hr As Long = 1
Public Const mc As Long = 1

'# Standard Formmating
Public Const fBoolean As String = "@" 'BOOLEAN format uses BOOLEAN datatype#
Public Const fCurrency As String = "_-""R$"" * #,##0.00_-;-""R$"" * #,##0.00_-" '#CURRENCY format uses DOUBLE datatype# 'NOTE: I had to change the dollar sign for the text "R$", and as my microsoft office is in portuguese, I also had to replace decimal dots by commas and vice versa
Public Const fDate As String = "dd/mm/yyyy" '#CAUTION!!! DATE format uses DOUBLE datatype for .value2 and DATE datatype for .value!!! In this model I'll use .value2 as much as I can, so dates will be DOUBLE variables#
Public Const fDouble As String = "0.00" '#DOUBLE format uses DOUBLE datatype#
Public Const fInteger As String = "0" '#INTEGER format uses DOUBLE datatype#
Public Const fLong As String = "0" '#LONG format uses DOUBLE datatype#
Public Const fString As String = "@" '#STRING format uses STRING datatype#
Public Const fPercentage As String = "0.00%" '#PERCENTAGE format uses DOUBLE datatype#

'####################################################################

'## The functions in this module are general functions

Public Function SheetExists(SheetName As String) As Boolean
    
Dim Sht As Excel.Worksheet
SheetExists = False
            
For Each Sht In Worksheets
    If SheetName = Sht.Name Then
        SheetExists = True
        Exit Function
    End If
Next Sht
End Function

Public Sub DeleteSheet(nWs As String)
'# Deletes a sheet without raising alerts
'# nWs = Worksheet Name
Application.DisplayAlerts = False
On Error Resume Next
ThisWorkbook.Worksheets(nWs).Delete
On Error GoTo 0
Application.DisplayAlerts = True


'''# Previously this was the function I used:
'''# Checks if a sheet exists and delete it without raising alerts
'''# nWs = Worksheet Name
''Dim Sht As Excel.Worksheet
''For Each Sht In Worksheets
''    If Sht.Name = nWs Then
''        Application.DisplayAlerts = False
''        ThisWorkbook.Worksheets(nWs).Delete
''        Application.DisplayAlerts = True
''        Exit For
''    End If
''Next Sht
End Sub

Public Sub DeleteSheetActivateOther(nwsToClose As String, nwsToActivate As String)
'# This function was created to use in buttons to Close sheets, as
'# you can't pass the command to activate another sheet after another function.
'# Buttons can only be associated with one macro at a time.
'# This functions allows more versatility, as can choose the sheet I want to close and activate

'# Be careful if you use this function inside your macros and not only in "Close" Buttons,
'# as activating another sheet CHANGES THE Selection Or Range!

'# Deletes a sheet without raising alerts and activates another sheet
'# nwsToClose = Name of the worksheet you want to close
'# nwsToActivate = Name of the worksheet you want to activate

Call DeleteSheet(nwsToClose)
ThisWorkbook.Worksheets(nwsToActivate).Activate
End Sub

Public Sub AddMacroButton(ByVal BtnR1 As Long, ByVal BtnC1 As Long, ByVal BtnName As String, ByVal OnActionCmd As String, ByVal WsName As String, Optional ByVal BtnR2 As Long, Optional ByVal BtnC2 As Long, Optional ByVal BtnWidth As Double, Optional ByVal BtnHeight As Double)
'This function is used to create a button in a worksheet
'It can use a single cell reference and a button width and height, OR two cells whose range will be the width and height of the button (in this case, button size may be affected by sheet formatting)

''Parameters:

'Range where button will be created
'BtnR1 - Row of Cel1 in range where you want to put the button
'BtnC1 - Column of Cell1 in range where you want to put the button

'BtnR2 - (OPTIONAL)(USE TOGETHER WITH BtnC2)(DO NOT USE IF USING BtnWidth AND BtnHeight) Row of Cell2 in range where you want to put the button.
'BtnC2 - (OPTIONAL)(USE TOGETHER WITH BtnR2)(DO NOT USE IF USING BtnWidth AND BtnHeight) Column of Cell2 in range where you want to put the button.

'BtnName - Name of the button for the users

'WsName - String with the name of the worksheet where the button will be created

'BtnWidth - (OPTIONAL)(USE TOGETHER WITH BtnHeight)(DO NOT USE IF USING BtnR2 AND BtnC2) Button size width

'BtnHeight - (OPTIONAL)(USE TOGETHER WITH BtnWidth)(DO NOT USE IF USING BtnR2 AND BtnC2) Button size height

''OnActionCmd - String with the command to be executed by the .OnAction property.
'               If you right-click a button, you can see the "Assign Macro" menu. The string in OnActionCmd will
'               be converted in the text shown there.
'               When using macros (sub or fuctions) which use one or more parameters, the whole string must be
'               given with inner single quotes around it. Take care where to use double or single quotes.
'               Macro without parameter: OnActionCmd = "Macro"
'               Macro with string parameter: OnActionCmd = "'Macro "String Parameter"'"    (As the parameter is a string, it NEEDS double quotes around it)
'               Macro with number parameter: OnActionCmd = "'Macro Number Parameter'"      (As the parameter is a number, it DOESN'T NEED double quotes around it)
'               Macro with mixed parameters: OnActionCmd = "'Macro Number Parameter,"String Parameter"'"
'               Look at the examples below:
'
'Macros can be passed writing the FULL or SHORT pathway name
'FULL - ModuleName.SubOrFunctionName + [PARAMETERS when needed]
'SHORT - SubOrFunctionName + [PARAMETERS when needed]
'
'Example passing a macro that does not require paramaters:
'Consider you have Sub MyMacro()
'If you pass argument "MyMacro" to OnActionCmd,
'Macro button assignment will show: 'Filename.Extension'!MyMacro  (note in this case the macro does not have single quotes around it)
'
'Example passing macro with string parameters using a FIXED string
'*By FIXED, I mean calling AddMacroButton will not need to use variables to pass as arguments to OnActionCmd
'Consider you have Sub MyMacro(ParamA as String)
'If you pass argument "'MyMacro "ABC"'" OR "'MyMacro """ & "ABC" & """'" to OnActionCmd,
'Macro button assignment will show: 'Filename.Extension'!'MyMacro "ABC"'
'
'Example passing macro with integer parameters using a FIXED integer
'*By FIXED, I mean calling AddMacroButton will not need to use variables to pass as arguments to OnActionCmd
'Consider you have Sub MyMacro(Param1 as Integer)
'If you pass argument "'MyMacro 123'" OR "'MyMacro " & 123 & "'" to OnActionCmd,
'Macro button assignment will show: 'Filename.Extension'!'MyMacro 123'
'
'Example passing macro with integer and string parameters using a FIXED integer and a FIXED string
'*By FIXED, I mean calling AddMacroButton will not need to use variables to pass as arguments to OnActionCmd
'Consider you have Sub MyMacro(Param1 as Integer, ParamA as String)
'If you pass argument "'MyMacro 123, "ABC"'" OR "'MyMacro " & 123 & ", """ & "ABC" & """'" to OnActionCmd,      (Note: the space after the comma is not necessary, but calling functions with more than one parameter normally is done this way)
'Macro button assignment will show: 'Filename.Extension'!'MyMacro 123, "ABC"'
'
'Example passing macro with string parameters using a string VARIABLE
'Consider you have Sub MyMacro(ParamA as String) and want to pass ParamA using the StrA variable (which has "ABC" value)
'If you pass argument "'MyMacro """ & StrA & """'" to OnActionCmd,
'Macro button assignment will show: 'Filename.Extension'!'MyMacro "ABC"'
'
'Example passing macro with integer parameters using an integer VARIABLE
'Consider you have Sub MyMacro(Param1 as Integer) and want to pass Param1 using the Int1 variable (which has 123 value)
'If you pass argument "'MyMacro " & Int1 & "'" to OnActionCmd,
'Macro button assignment will show: 'Filename.Extension'!'MyMacro 123'
'
'Example passing macro with integer and string parameters using an integer VARIABLE and a string VARIABLE
'Consider you have Sub MyMacro(Param1 as Integer, ParamA as String) and want to pass Param1 using the Int1 variable (which has 123 value) and pass ParamA using the StrA variable (which has "ABC" value)
'If you pass argument "'MyMacro " & Int1 & ", """ & StrA & """'" to OnActionCmd,      (Note: the space after the comma is not necessary, but calling functions with more than one parameter normally is done this way)
'Macro button assignment will show: 'Filename.Extension'!'MyMacro 123, "ABC"'

'Asserting both BtnR2 and BtnC2 have been given
If BtnR2 <> 0 And BtnC2 = 0 Then
    BtnR2 = 0
ElseIf BtnR2 = 0 And BtnC2 <> 0 Then
    BtnC2 = 0
End If

'Asserting both BtnWidth and BtnHeight have been given
If BtnWidth <> 0 And BtnHeight = 0 Then
    BtnWidth = 0
ElseIf BtnWidth = 0 And BtnHeight <> 0 Then
    BtnHeight = 0
End If

'Asserting BtnWidth and BtnHeight will be used when given
If BtnWidth <> 0 And BtnHeight <> 0 And BtnR2 <> 0 And BtnC2 <> 0 Then
    BtnR2 = BtnR1
    BtnC2 = BtnC1
End If

'This is to write less code below
If (BtnR2 = 0 And BtnC2 = 0) Then
    BtnR2 = BtnR1
    BtnC2 = BtnC1
End If

'Button creation
Dim Btn As Button
Dim BtnRng As Range
Set BtnRng = ThisWorkbook.Worksheets(WsName).Range(Cells(BtnR1, BtnC1), Cells(BtnR2, BtnC2))

'Attention! If you use BtnRng.Width and BtnRng.Height, the button size may be affected by the sheet's formatting
If BtnWidth = 0 And BtnHeight = 0 Then
    Set Btn = ThisWorkbook.Worksheets(WsName).Buttons.Add(BtnRng.Left, BtnRng.Top, BtnRng.Width, BtnRng.Height)
Else
    Set Btn = ThisWorkbook.Worksheets(WsName).Buttons.Add(BtnRng.Left, BtnRng.Top, BtnWidth, BtnHeight)
End If

'Button configuration
With Btn
    .Caption = BtnName
    .OnAction = OnActionCmd
End With

End Sub

Public Sub DeactivateSystemFunctions()
'# This code is used at the beggining of macros to speed up the processing,
'# as it deactivates functions that delay code execution.
'# Remember to reactivate these functions in the end of the
'# macro using Sub ReactivateSystemFunctions()
With Application
    .ScreenUpdating = False
    .DisplayStatusBar = False
    .Calculation = xlCalculationManual
    .ActiveSheet.DisplayPageBreaks = False
    .EnableEvents = False
End With

End Sub

Public Sub ReactivateSystemFunctions()
'# This code is used at the end of macros, to reactivate some of Excel's standard functions
'# that have been deactivated by Sub DeactivateSystemFunctions().
With Application
    .ScreenUpdating = True
    .DisplayStatusBar = True
    .Calculation = xlCalculationAutomatic
    .ActiveSheet.DisplayPageBreaks = True
    .EnableEvents = True
End With

End Sub

Function ColumnLetter(ByVal ColumnNumber As Long) As String
'# This functions converts a column number to its correspondent letter
'# Got this macro from https://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
'# ColNumber = Column Number
    Dim arr
    arr = Split(Cells(1, ColumnNumber).Address(True, False), "$")
    ColumnLetter = arr(0)
End Function

Function ArrayFilteredRange(rngFiltered As Range, IncludeHeader As Boolean) As Variant
'# Makes an array from a filtered range. A filtered range has non contiguous
'# rows, so you can't transfer that range directly into an array.

'# After Filtering (or filtering and sorting) the worksheet,
'# pass the filtered range to this function as the example below:
'# Set rngFiltered = wsTemp.Range(SET WHOLE RANGE USED BEFORE FILTERING,
'# THERE IS NO NEED TO FIND THE RESULTING FILTERED RANGE).SpecialCells(xlCellTypeVisible)


'# After filtering (or filtering and sorting) the worksheet,
'# test if the filter returned results, as in the example below:
'frExample = hrExample
'lrExample = wsExample.Cells(Rows.Count, cIdExample).End(xlUp).Row
'fcExample = mcExample
'lcExample = wsExample.Cells(hrExample, Columns.Count).End(xlToLeft).Column
'If frExample = lrExample Then
'    MsgBox ("There are no results to show." & vbCrLf & "Action cancelled.")
'    wsExample.Activate
'    Call DeleteSheet(nwsExample)
'    Call ReactivateSystemFunctions
'    Exit Sub
'End If

'# Then, if there are results after the filter, call this
'# function using rng.SpecialCells(xlCellTypeVisible), as
'# in the example below:
'Set rng = wsExample.Range(wsExample.Cells(frExample, fcExample), wsExample.Cells(lrExample, lcExample))
'arrTemp = ArrayFilteredRange(rng.SpecialCells(xlCellTypeVisible), False)
'Set rng = Nothing

'# adapted from Ambie's answer in Aug 26 '17 at stackoverflow:
'# https://stackoverflow.com/questions/45897753/get-a-filtered-range-into-an-array

'Dim rngFiltered As Range
Dim myArea As Range, myRow As Range, myCell As Range
Dim arrFiltered() As Variant
Dim r As Long

If IncludeHeader = False Then
    With rngFiltered
        r = -1 '# Start at -1 to remove header row
        For Each myArea In rngFiltered.Areas
            r = r + myArea.Rows.Count
        Next
        ReDim arrFiltered(1 To r, 1 To .Columns.Count)
    End With
    r = 1
    For Each myArea In rngFiltered.Areas
        For Each myRow In myArea.Rows
            If myRow.Row <> 1 Then '# This if statement removes the header row
                For Each myCell In myRow.Cells
                    arrFiltered(r, myCell.Column) = myCell.Value2
                Next
                r = r + 1
            End If
        Next
    Next

ElseIf IncludeHeader = True Then
    With rngFiltered
        r = 0 '# Start at 0 to include header row
        For Each myArea In rngFiltered.Areas
            r = r + myArea.Rows.Count
        Next
        ReDim arrFiltered(1 To r, 1 To .Columns.Count)
    End With
    r = 1
    For Each myArea In rngFiltered.Areas
        For Each myRow In myArea.Rows
            'If myRow.row <> 1 Then '# This if statement removes the header row
                For Each myCell In myRow.Cells
                    arrFiltered(r, myCell.Column) = myCell.Value2
                Next
                r = r + 1
            'End If
        Next
    Next
End If
ArrayFilteredRange = arrFiltered
End Function

Public Sub Bubblesort(ByRef ArrayToSort(), ByVal Lower As Long, ByVal Upper As Long, Optional ByVal TextCompare As Boolean = False)
'# Sorts an array using Bubblesort
'# ATTENTION! While comparing strings with TextCompare = False, the comparison will be CASE SENSITIVE (Binary Comparison).

'# ArrayToSort() - Array to sort
'# Lower - The lower element position to use in the sorting
'# Upper - The upper element position to use in the sorting

'# TextCompare - If True, will compare elements using StrComp(A,B, CompareMethod)
'# CompareMethod will be 1, but here are other values:
'# If value =  1, uses Text Compare (CASE INSENSITIVE);
'# If value =  0, uses Binary Compare (CASE SENSITIVE);
'# If value = -1, when using 'Option Compare XXXX' at the top of the module, it will use that kind of comparison (example: If using 'Option Compare Text' declaration in the beggining of the module, it will perform a Text Compare, as if it's value was 1)
'# Result values for StrComp function:
'# -1, if StrComp(A<B)
'# 0, if StrComp(A=B)
'# 1, if StrComp(A>B)

Dim ElementA As Long
Dim ElementB As Long
Dim TempValue As Variant

'# Binary compare
If TextCompare = False Then
    For ElementA = 0 To (UBound(ArrayToSort()) - 1)
        For ElementB = ElementA + 1 To UBound(ArrayToSort())
            If ArrayToSort(ElementA) > ArrayToSort(ElementB) Then
                TempValue = ArrayToSort(ElementA)
                ArrayToSort(ElementA) = ArrayToSort(ElementB)
                ArrayToSort(ElementB) = TempValue
            End If
        Next ElementB
    Next ElementA

'# Text compare
ElseIf TextCompare = True Then
    For ElementA = 0 To (UBound(ArrayToSort()) - 1)
        For ElementB = ElementA + 1 To UBound(ArrayToSort())
            If StrComp(ArrayToSort(ElementA), ArrayToSort(ElementB), 1) = 1 Then
                TempValue = ArrayToSort(ElementA)
                ArrayToSort(ElementA) = ArrayToSort(ElementB)
                ArrayToSort(ElementB) = TempValue
            End If
        Next ElementB
    Next ElementA
End If
End Sub

Sub Quicksort(ByRef values(), ByVal min As Long, ByVal max As Long)

Dim med_value As Variant
Dim Hi As Long
Dim Lo As Long
Dim i As Long

' If the list has only 1 item, it's sorted.
If min >= max Then
    Exit Sub
End If

' Pick a dividing item randomly.
i = min + Int(Rnd(max - min + 1))
med_value = values(i)

' Swap the dividing item to the front of the list.
values(i) = values(min)

' Separate the list into sublists.
Lo = min
Hi = max

Do
    ' Look down from hi for a value < med_value.
    Do While values(Hi) >= med_value
        Hi = Hi - 1
        If Hi <= Lo Then
            Exit Do
        End If
    Loop

    If Hi <= Lo Then
        ' The list is separated.
        values(Lo) = med_value
        Exit Do
    End If

    ' Swap the lo and hi values.
    values(Lo) = values(Hi)

    ' Look up from lo for a value >= med_value.
    Lo = Lo + 1
  
    Do While values(Lo) < med_value
        Lo = Lo + 1
        If Lo >= Hi Then
            Exit Do
        End If
    Loop
  
    If Lo >= Hi Then
        ' The list is separated.
        Lo = Hi
        values(Hi) = med_value
        Exit Do
    End If

    ' Swap the lo and hi values.
    values(Hi) = values(Lo)
  
Loop ' Loop until the list is separated.

' Recursively sort the sublists.
Call Quicksort(values, min, Lo - 1)
Call Quicksort(values, Lo + 1, max)
End Sub


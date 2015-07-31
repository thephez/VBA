Option Explicit
'Inputs/Externals
Dim Door_Open_Input, DPS_Input, Door_Open_Output, Emergency_Release, Shunt_Input, Card_Reader_H_Input, Card_Reader_H_Enable, Hold_Unlock_Input, Card_Reader_S_Input As Integer
Dim Card_Reader_S_Enable, Exit_Push_Button, Future, Half_Cycle_Lock, Future4, OTL_Output, BRC_Output, IOStartRow, InputNumOfObjects, OutputNumOfObjects, StartRow As Integer
Dim Door_Open_Input_Slider, Door_Stop_Input_Slider, Door_Close_Input_Slider, Door_Shunt_Input_Slider As Integer
Dim Door_OTL_Output_Slider, hVehicle, hKey, hPanic  'Floating Point
Dim ooo



Sub Alarms()
Dim Object, ObjectCol, ObjectSelectCol, NumOfObjects, Alarm, AlarmCol, AlarmSelectCol, NumOfAlarms, i, j, k, l, StartRow, AlarmsCol As Integer
Dim TagName, cType, Limit, AlarmMessage, Priority, Selection As Integer
StartRow = 2    'The 1st row is just the type of data. The second row is where the actual data is.
ObjectCol = 1   '"A"
ObjectSelectCol = ObjectCol + 1 '"B"
AlarmCol = ObjectSelectCol + 1    '"C"
AlarmSelectCol = AlarmCol + 1  '"D"
'AlarmsCol = 10 '"J", Where the actual results will print
TagName = AlarmSelectCol + 8    'Give the input columns 3 column spaces from output columns
cType = TagName + 1
Limit = cType + 1
AlarmMessage = Limit + 1
Priority = AlarmMessage + 1
Selection = Priority + 1
NumOfObjects = Enumerate(StartRow, ObjectCol)
NumOfAlarms = Enumerate(StartRow, AlarmCol)
k = StartRow    'k denotes the row of the output
For i = StartRow To NumOfObjects + StartRow - 1
    If Cells(i, ObjectSelectCol) <> "" Then
    For l = 1 To Cells(i, ObjectSelectCol)
        For j = StartRow To NumOfAlarms + StartRow - 1
            If Cells(j, AlarmSelectCol) <> "" Then
                Cells(k, TagName) = Cells(i, ObjectCol) & "[" & l & "]" & "." & Cells(j, AlarmCol)
                Cells(k, cType) = Cells(j, 5)
                Cells(k, Limit) = Cells(j, 6)
                If Cells(i, ObjectCol) = "Redundancy_HMI" Then
                    Cells(k, AlarmMessage) = "{CompName[" & l & "]} " & Cells(j, 7)
                Else
                    Cells(k, AlarmMessage) = "{" & Cells(i, ObjectCol) & "[" & l & "].strName} " & Cells(j, 7)
                End If
                Cells(k, Priority) = Cells(j, 8)
                Cells(k, Selection) = Cells(j, 9)
                k = k + 1
            End If
        Next
    Next
    End If
Next
ActiveSheet.Range(Cells(StartRow, TagName), Cells(k - 1, Selection)).Select
End Sub

Public Function Enumerate(StartRow, Column)
Dim i, Tally
Tally = 0
For i = StartRow To 3000
    If Cells(i, Column) <> "" Then Tally = Tally + 1
Next
'Cells(1, 4) = Tally
Enumerate = Tally
End Function

Public Sub ClearFieldsAlarms()
Dim ObjectSelectCol, NumOfObjects, AlarmSelectCol, NumOfAlarms, StartRow, i As Integer
Dim ObjectCol, AlarmCol As Integer
StartRow = 2    'The 1st row is just the type of data. The second row is where the actual data is.
ObjectSelectCol = 2 '"B"
AlarmSelectCol = 4  '"D"
ObjectCol = 1   '"A"
AlarmCol = 3    '"C"
NumOfObjects = Enumerate(StartRow, ObjectCol)
NumOfAlarms = Enumerate(StartRow, AlarmCol)
For i = StartRow To NumOfObjects + StartRow - 1
    Cells(i, ObjectSelectCol) = ""
Next
For i = StartRow To NumOfAlarms + StartRow - 1
    Cells(i, AlarmSelectCol) = ""
Next
End Sub

Public Sub ClearFieldsSort()
Dim NumOfObjects, AlarmSelectCol, NumOfAlarms, StartRow, i As Integer
Dim ObjectCol, GroupActiveCol, AlarmCol, IndexCol As Integer
StartRow = 2    'The 1st row is just the type of data. The second row is where the actual data is.
ObjectCol = 1   '"A"
GroupActiveCol = ObjectCol + 1
IndexCol = GroupActiveCol + 1
AlarmCol = IndexCol + 2
AlarmSelectCol = AlarmCol + 1
NumOfObjects = Enumerate(StartRow, ObjectCol)
NumOfAlarms = Enumerate(StartRow, AlarmCol)
For i = StartRow To NumOfObjects + StartRow - 1
    Cells(i, ObjectCol) = ""
    Cells(i, GroupActiveCol) = ""
    Cells(i, IndexCol) = ""
Next
For i = StartRow To NumOfAlarms + StartRow - 1
    Cells(i, AlarmSelectCol) = ""
Next
End Sub

Public Sub ClearListAlarms()
Dim i, j, StartCol, StartRow, NumOfEntries, EndColumn
StartRow = 2
EndColumn = 8
For i = 1 To 30
    If Cells(1, i) = "Tag Name" Then StartCol = i
Next
ActiveSheet.Cells.Range(Cells(StartRow, 12), Cells(3000, 17)).ClearContents
End Sub

Public Sub ClearListTags()
Dim i, j, StartCol, StartRow, NumOfEntries, EndColumn
StartRow = 2
EndColumn = 8
For i = 1 To 30
    If Cells(1, i) = "Tag Name" Then StartCol = i
Next
ActiveSheet.Cells.Range(Cells(StartRow, 12), Cells(3000, 17)).ClearContents
End Sub

Public Sub ClearListSort()
Dim i, j, StartCol, StartRow, NumOfEntries, EndColumn
StartRow = 1
EndColumn = 8
For i = 1 To 30
    If Cells(1, i) = "HMI[1] Alarms" Then StartCol = i
Next
ActiveSheet.Cells.Range(Cells(StartRow, 15), Cells(200, 15500)).ClearContents
End Sub

Public Sub TagList()
Dim i, j, FromIndex, ToIndex, Object, Addendum, Selection, Tags, StartRow, start, ending, SelectedRow As Integer
StartRow = 2
Object = 1  '"A"
Addendum = Object + 1
Selection = Addendum + 2
FromIndex = Selection + 1
ToIndex = FromIndex + 1
Tags = ToIndex + 3
start = Cells(StartRow, FromIndex)
ending = Cells(StartRow, ToIndex)
SelectedRow = Cells(StartRow, Selection)
j = StartRow
For i = start To ending
    Cells(j, Tags) = Cells(SelectedRow, Object) & "[" & i & "]" & Cells(StartRow, Addendum)
    j = j + 1
Next
ActiveSheet.Range(Cells(StartRow, Tags), Cells(j - 1, Tags)).Select
End Sub
Public Sub Sort()
Dim Object, GroupActive, i, j As Integer
Dim Char, number
Object = 1
GroupActive = Object + 1
Call Indexes
Call PrepareColumns
End Sub
Public Function Indexes()   'Called by Sort()
Dim Char, number, i, j, GroupActive, NumOfObjects, StartRow, ObjectCol
StartRow = 2
ObjectCol = 1
GroupActive = ObjectCol + 1
Char = ""
number = ""
NumOfObjects = Enumerate(StartRow, ObjectCol)
For i = StartRow To NumOfObjects + 1
    For j = 0 To 3
    Char = Mid(Cells(i, GroupActive), 13 + j, 1)
        If Char <> "]" Then 'IsNumeric(Char) Then
            number = number & Char
        End If
    Next
    Cells(i, 3) = number
    number = ""
Next
End Function

Public Function GetLowest(StartRow, IndexCol)
Dim i, NumOfObjects, Lowest, Last
Lowest = 1000
Last = 0
NumOfObjects = Enumerate(StartRow, IndexCol) + 1
For i = StartRow To NumOfObjects
    If Cells(i, IndexCol) < Last Then
        Lowest = Cells(i, IndexCol)
    End If
    Last = Cells(i, IndexCol)
Next
Cells(NumOfObjects + 1, IndexCol) = "Lowest Index: " & Lowest
End Function

Public Function GetHighest(StartRow, IndexCol)
Dim i, NumOfObjects, Highest, Last
Highest = 0
Last = 0
NumOfObjects = Enumerate(StartRow, IndexCol) + 1
For i = StartRow To NumOfObjects
    If Cells(i, IndexCol) > Last Then
        Highest = Cells(i, IndexCol)
    End If
    Last = Cells(i, IndexCol)
Next
Cells(NumOfObjects + 2, IndexCol) = "Highest Index: " & Highest
End Function

Public Sub PrepareColumns() 'Called by Sort()
Dim i, j, k, GroupActive, NumOfObjects, StartRow, ObjectCol, index, IndexCol, LastCol, AlarmCol, AlarmSelectCol, NumOfAlarms, HighestIndex, LowestIndex, BlankMode As Integer
Dim TypeCol, LimitCol, MessageCol, PriorityCol, SelectionCol, GroupActiveCol, ColumnOffset, ResultsStartCol, Length As Integer
Dim ObjectType As String
Dim got(300) As Integer    'This is an array that will hold the indexes encountered in numerical order. If index = 4, then got(index) is incremented by 1
StartRow = 2    'The row just beneath the headers
ObjectCol = 1   'Column number of the PDoor[n], OSCDoor[n], Light[n], etc copy/pasted in by the user
GroupActiveCol = ObjectCol + 1  'Column number where GroupActives[n] are copy/pasted in by the user
IndexCol = GroupActiveCol + 1    'The column where Indexes() has printed the indexes extracted from "GroupActive[n]"
AlarmCol = IndexCol + 2 'Indusoft "Alarm" column. Leaving a space between alarms and index to separate user input from Indusoft Alarm parameters section
AlarmSelectCol = AlarmCol + 1   'Column where the user puts a 1 next to the alarm desired
TypeCol = AlarmSelectCol + 1  'Indusoft "type" column ('Hi' or 'HiHi')
LimitCol = TypeCol + 1  'Indusoft "Limit" column - The value that triggers the alert in Indusoft
MessageCol = LimitCol + 1   'Indusoft "Message" column - The corresponding message
PriorityCol = MessageCol + 1    'Indusoft "Priority" column - Prioritizes alarms
SelectionCol = PriorityCol + 1  'Indusoft "Selection" column - not sure how Indusoft uses this...
LastCol = SelectionCol  'Last alarm parameters column. After this, the sorted output can begin to the right.
ResultsStartCol = LastCol + 5   'The starting column where the results will be printed
BlankMode = 1   'Used as a Boolean to determine whether any alarm has been selected - 1 when no alarm selected
ColumnOffset = 0    'This is how many columns should separate each sorted group of alarms in the results area. Should be 1 if BlankMode, 7-8 otherwise
NumOfObjects = Enumerate(StartRow, ObjectCol)  'Number of of Objects pasted in by the user.
NumOfAlarms = Enumerate(StartRow, AlarmCol) + 1 'Number of Alarms in the alarms column
For i = StartRow To NumOfAlarms 'Loop Through the AlarmSelectedCol to see if any alarms selected
    If Cells(i, AlarmSelectCol) <> "" Then  'If any cell contains something, Reset BlankMode
        BlankMode = 0
    End If
Next
If BlankMode Then ColumnOffset = 3 Else ColumnOffset = LastCol - AlarmCol  'The number of columns encompassed by Indusoft Alarm parameters
Call ClearListSort  'Clear the last results on the spreadsheet
j = 0   'index used to loop through the items in AlarmSelectCol, that is, if BlankMode is not set
For i = StartRow To NumOfObjects + 1    'Cycling through the objects
    index = Cells(i, IndexCol)
    If index <> 0 Or index <> "" Then   'if index is not blank or 0
        Length = Len(Cells(i, ObjectCol)) - 5
        ObjectType = Left(Cells(i, ObjectCol), Length)
        'If index > 8 Then index = 1
        If BlankMode Then ResultsStartCol = LastCol + index * ColumnOffset + 5 Else ResultsStartCol = LastCol + index * ColumnOffset
        If got(index) = 0 Then    'If you've not yet gotten an index = '1' - that is, you've gotten this index 0 times so far
            Cells(1, ResultsStartCol) = "GroupActive[" & index & "] Alarms" 'Title
            got(index) = got(index) + 1
        End If
        If BlankMode Then   'If there are no alarms selected in the AlarmSelectCol
            got(index) = got(index) + 1
            Cells(got(index), ResultsStartCol) = Cells(i, ObjectCol)  'Only print the Object in the OjectCol
        Else:   'otherwise if there are alarms selected
            For j = StartRow To NumOfAlarms 'Cycling through the alarmSelectedCol's items
                If Cells(j, AlarmSelectCol) <> "" Then
                    got(index) = got(index) + 1
                    Cells(got(index), ResultsStartCol) = Cells(i, ObjectCol) & "." & Cells(j, AlarmCol)
                    Cells(got(index), (ResultsStartCol + 1)) = Cells(j, TypeCol)
                    Cells(got(index), (ResultsStartCol + 2)) = Cells(j, LimitCol)
                    Cells(got(index), (ResultsStartCol + 3)) = "{" & Cells(i, ObjectCol) & ".strName}" & Cells(j, MessageCol)
                    'Cells(got(index), (ResultsStartCol + 3)) = "{" & Cells(i, ObjectCol) & MemberName(ObjectType) & "}" & Cells(j, MessageCol)
                    Cells(got(index), (ResultsStartCol + 4)) = Cells(j, PriorityCol)
                    Cells(got(index), (ResultsStartCol + 5)) = Cells(j, SelectionCol)
                    'Cells(got(index), (ResultsStartCol + 6)) = "       "
                End If
            Next
        End If
        Cells(1, ResultsStartCol + 1) = got(index) - 1 & " item(s)"
        ResultsStartCol = 0
    End If
Next
Dim TargetRange As Range
Dim cel As Range
Set TargetRange = Range(Cells(1, LastCol + 5), Cells(300, 1000))
TargetRange.Columns.AutoFit  'Fit the cells to the text
HighestIndex = Enumerate(StartRow, IndexCol)
For i = (LastCol + 5) To 300 '(lastCol + 5 + HighestIndex * ColumnOffset)
    If Cells(2, i) = Empty Then Cells(2, i).EntireColumn.Hidden = True
Next
End Sub

Public Function MemberName(Object) As String    'Called in PrepareColumns
Select Case (Object)
    Case "Pdoor", "OSCDoor", "Mdoor": MemberName = ".strName"
    Case "UPS_Alarm": MemberName = ".intFloor"
End Select
End Function

Public Sub PhraseBuilder()
Dim i, j, FromIndex, ToIndex, Beginning, ending, Phrase, StartRow, start As Integer
StartRow = 2
Beginning = 1  '"A"
ending = Beginning + 1
FromIndex = ending + 1
ToIndex = FromIndex + 1
Phrase = ToIndex + 3
FromIndex = Cells(StartRow, FromIndex)
ToIndex = Cells(StartRow, ToIndex)
j = StartRow
For i = FromIndex To ToIndex
    Cells(j, Phrase) = Cells(StartRow, Beginning) & i & Cells(StartRow, ending)
    j = j + 1
Next
ActiveSheet.Range(Cells(StartRow, Phrase), Cells(j - 1, Phrase)).Select
End Sub

Public Sub PLCDoorMnemonic()
Dim i, j, k, m, n, start, NumOfObjects, StartRow, DoorNameCol, IndexCol, DPSInputCol, DPSWCol, LOKCol, LOKWCol, LOKOutputCol As Integer
Dim MnemonicCol, AddressCol, RackLocationCol, UsageCol, CommentCol, NamesCol, DataTypeCol, HighestIndex, DriverCol1, DriverCol2, DriverCol3, DriverCol4, StartAt As Integer
StartRow = 2    'The row just beneath the headers
DoorNameCol = 1   'Column number of the DoorNames Column
IndexCol = DoorNameCol + 1    'IndusoftIndex Column
DPSInputCol = IndexCol + 1    'PLC address of the DPS bit Column - Door_Open_Input Ex: I2A101_DPS 0.00
DPSWCol = DPSInputCol + 1   'PLC W bit for doors - DPS_Input Ex: W2A101_DPS W25.00
LOKWCol = DPSWCol + 1 'PLC address of the LOK bit Column - Shunt_Input Ex: WO2A101_LOK W50.00
LOKOutputCol = LOKWCol + 1  'Ouput LOK bit Ex: O2A101_LOK 100.00
MnemonicCol = LOKOutputCol + 2
DriverCol1 = MnemonicCol + 2
DriverCol2 = DriverCol1 + 1
DriverCol3 = DriverCol2 + 2
DriverCol4 = DriverCol3 + 1
NamesCol = DriverCol4 + 2
DataTypeCol = NamesCol + 1
AddressCol = DataTypeCol + 1
CommentCol = AddressCol + 1
StartAt = MnemonicCol + 1
NumOfObjects = Enumerate(StartRow, DoorNameCol)  'Number of of Objects pasted in by the user.
'Call ClearListSort  'Clear the last results on the spreadsheet
j = 2   'index for the Mnemonics
k = 2   'index for the Names and address
m = 2   'Index for the LOK Indusoft Driver sheets - gets its own index b/c not every door has a LOK
n = 2   'Index for the DPS Indusoft Driver Sheets
start = Cells(2, StartAt)
For i = StartRow To 3000    'Cycling through the doors
If Cells(i, DoorNameCol) <> "" Then
    'Output Mnemonics
    Cells(j, MnemonicCol) = " ' " & "Door " & start & " - " & Cells(i, DoorNameCol) 'Rung comment section
    j = j + 1
    Cells(j, MnemonicCol) = "LD " & FormatNumber(Cells(i, DPSInputCol), 2)
    j = j + 1
    'Only if Factory Test PDoors needs DPS simulation
    'If Cells(i, LOKOutputCol) <> "" Then    'If this is a pdoor
    '    Cells(j, MnemonicCol) = "LD H511.15"
    '    j = j + 1
    '    Cells(j, MnemonicCol) = "ANDNOT W" & FormatNumber(Cells(i, LOKWCol), 2)
     '   j = j + 1
    '    Cells(j, MnemonicCol) = "ORLD"
    '    j = j + 1
    'End If
    Cells(j, MnemonicCol) = "OUT W" & FormatNumber(Cells(i, DPSWCol), 2)
    j = j + 1
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(j, MnemonicCol) = "LD W" & FormatNumber(Cells(i, LOKWCol), 2)
        j = j + 1
        Cells(j, MnemonicCol) = "OUT " & FormatNumber(Cells(i, LOKOutputCol), 2)
        j = j + 1
    End If
    start = start + 1
    'Output Indusoft DPS Driver sheet elements
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(n, DriverCol1) = "PDoor_A[" & Cells(i, IndexCol) & "].bolDPS"
    Else: Cells(n, DriverCol1) = "MDoor_A[" & Cells(i, IndexCol) & "].bolDPS"
    End If
    Cells(n, DriverCol2) = Cells(i, DPSInputCol) + 25
    n = n + 1
    'Output Indusoft LOK Driver sheet elements
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(m, DriverCol3) = "PDoor_A[" & Cells(i, IndexCol) & "].bolLOK"
        Cells(m, DriverCol4) = Cells(i, DPSInputCol)
        m = m + 1
    End If
    'Output names and addresses for CX Programmer Symbols Page
    Cells(k, NamesCol) = "I" & Cells(i, DoorNameCol) & "_DPS" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = FormatNumber(Cells(i, DPSInputCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "DPS Input" 'PLC bit Address
    k = k + 1
    Cells(k, NamesCol) = "W" & Cells(i, DoorNameCol) & "_DPS" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "W" & FormatNumber(Cells(i, DPSWCol), 2)     'Cells(i, DPSWCol) 'PLC bit Address
    Cells(k, CommentCol) = "DPS" 'PLC bit Address
    k = k + 1
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(k, NamesCol) = "WO" & Cells(i, DoorNameCol) & "_LOK" 'PLC bit name
        Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
        Cells(k, AddressCol) = "W" & FormatNumber(Cells(i, LOKWCol), 2) 'PLC bit Address
        Cells(k, CommentCol) = "LOK" 'PLC bit Address
        k = k + 1
    End If
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(k, NamesCol) = "O" & Cells(i, DoorNameCol) & "_LOK" 'PLC bit name
        Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
        Cells(k, AddressCol) = FormatNumber(Cells(i, LOKOutputCol), 2) 'PLC bit Address
        Cells(k, CommentCol) = "LOK Output" 'PLC bit Address
        k = k + 1
    End If
End If
Next
'Dim TargetRange As Range
'Dim cel As Range
'Set TargetRange = Range(Cells(1, lastCol + 5), Cells(300, 1000))
'TargetRange.Columns.AutoFit  'Fit the cells to the text
'HighestIndex = Enumerate(StartRow, IndexCol)
'For i = (lastCol + 5) To 300 '(lastCol + 5 + HighestIndex * ColumnOffset)
'    If Cells(2, i) = Empty Then Cells(2, i).EntireColumn.Hidden = True
'Next
End Sub

Public Sub PLCEMRMnemonic()
Dim i, j, k, m, n, l, NumOfObjects, StartRow, DoorNameCol, IndexCol, EMRCol, LOKOutputCol As Integer
Dim SHTCol, HCRDInputCol, HCRDEnableCol, HULInputCol, SCRDInputCol, SCRDEnableCol, ExitBtnCol, BRCOutputCol, OTLOutputCol, HCRDPhysColnput, SCRDPhysColnput As Integer
Dim MnemonicCol, AddressCol, CommentCol, NamesCol, DataTypeCol, HighestIndex, EMRCol1, EMRCol2 As Integer
Dim EMRDriverCol1, EMRDriverCol2, OTLDriverCol1, OTLDriverCol2, BRCDriverCol1, BRCDriverCol2, SHTDriverCol1, SHTDriverCol2, HCRDEnableDriverCol1, HCRDEnableDriverCol2, HULDriverCol1, HULDriverCol2, SCRDEnableDriverCol1, SCRDEnableDriverCol2 As Integer
StartRow = 2    'The row just beneath the headers
DoorNameCol = 1   'Column number of the DoorNames Column
IndexCol = DoorNameCol + 1    'IndusoftIndex Column
EMRCol = IndexCol + 1    'PLC address of the EMR bit Column - Door_Open_Input Ex: W2D209_EMR 75.00

SHTCol = EMRCol + 1
HCRDInputCol = SHTCol + 1
HCRDEnableCol = HCRDInputCol + 1
HULInputCol = HCRDEnableCol + 1
SCRDInputCol = HULInputCol + 1
SCRDEnableCol = SCRDInputCol + 1
ExitBtnCol = SCRDEnableCol + 1
BRCOutputCol = ExitBtnCol + 1
OTLOutputCol = BRCOutputCol + 1

LOKOutputCol = OTLOutputCol + 1
HCRDPhysColnput = LOKOutputCol + 1
SCRDPhysColnput = HCRDPhysColnput + 1
MnemonicCol = LOKOutputCol + 4
NamesCol = MnemonicCol + 2
DataTypeCol = NamesCol + 1
AddressCol = DataTypeCol + 1
CommentCol = AddressCol + 1

BRCDriverCol1 = CommentCol + 2
BRCDriverCol2 = BRCDriverCol1 + 1
OTLDriverCol1 = BRCDriverCol2 + 2
OTLDriverCol2 = OTLDriverCol1 + 1
EMRDriverCol1 = OTLDriverCol2 + 2
EMRDriverCol2 = EMRDriverCol1 + 1
SHTDriverCol1 = EMRDriverCol2 + 2
SHTDriverCol2 = SHTDriverCol1 + 1
HCRDEnableDriverCol1 = SHTDriverCol2 + 2
HCRDEnableDriverCol2 = HCRDEnableDriverCol1 + 1
HULDriverCol1 = HCRDEnableDriverCol2 + 2
HULDriverCol2 = HULDriverCol1 + 1
SCRDEnableDriverCol1 = HULDriverCol2 + 2
SCRDEnableDriverCol2 = SCRDEnableDriverCol1 + 1

NumOfObjects = Enumerate(StartRow, DoorNameCol)  'Number of of Objects pasted in by the user.
'Call ClearListSort  'Clear the last results on the spreadsheet
j = 2   'index for the Mnemonics
k = 2   'index for the Names and address
m = 2   'Index for the LOK Indusoft Driver sheets - gets its own index b/c not every door has a LOK
l = 1
n = 2
For i = StartRow To 3000    'Cycling through the doors
If Cells(i, DoorNameCol) <> "" Then
    'Output Mnemonics
    If Cells(i, HCRDPhysColnput) <> "" Then
        Cells(j, MnemonicCol) = " ' " & "Card Reader" & l & " - " & Cells(i, DoorNameCol) 'Rung comment section
        j = j + 1
        Cells(j, MnemonicCol) = "LD " & FormatNumber(Cells(i, HCRDPhysColnput), 2)
        j = j + 1
        Cells(j, MnemonicCol) = "AND W" & Cells(i, DoorNameCol) & "_HCRDEnable"
        j = j + 1
        Cells(j, MnemonicCol) = "OUT W" & Cells(i, DoorNameCol) & "_CRD_H_IN"
        j = j + 1
        l = l + 1
    End If
    If Cells(i, SCRDPhysColnput) <> "" Then
        Cells(j, MnemonicCol) = " ' " & "Card Reader" & l & " - " & Cells(i, DoorNameCol) 'Rung comment section
        j = j + 1
        Cells(j, MnemonicCol) = "LD " & FormatNumber(Cells(i, SCRDPhysColnput), 2)
        j = j + 1
        Cells(j, MnemonicCol) = "AND W" & Cells(i, DoorNameCol) & "_SCRDEnable"
        j = j + 1
        Cells(j, MnemonicCol) = "OUT W" & Cells(i, DoorNameCol) & "_CRD_S_IN"
        j = j + 1
        l = l + 1
    End If
    
    'Output Indusoft EMR Driver sheet elements
    'BRC
    If Cells(i, LOKOutputCol) <> "" Then    'If the door has a lok bit - in other words, if it is a PDoor
        Cells(n, BRCDriverCol1) = "PDoor_A[" & Cells(i, IndexCol) & "].bolBRC"
    Else: Cells(n, BRCDriverCol1) = "MDoor_A[" & Cells(i, IndexCol) & "].bolBRC"
    End If
    Cells(n, BRCDriverCol2) = FormatNumber(Cells(i, BRCOutputCol), 2)
    'OTL
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(n, OTLDriverCol1) = "PDoor_A[" & Cells(i, IndexCol) & "].bolOTL"
    Else: Cells(n, OTLDriverCol1) = "MDoor_A[" & Cells(i, IndexCol) & "].bolOTL"
    End If
    Cells(n, OTLDriverCol2) = FormatNumber(Cells(i, OTLOutputCol), 2)
    'EMR
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(n, EMRDriverCol1) = "PDoor_A[" & Cells(i, IndexCol) & "].bolEMR"
        Cells(n, EMRDriverCol2) = FormatNumber(Cells(i, EMRCol), 2)
    End If
    'SHT
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(n, SHTDriverCol1) = "PDoor_A[" & Cells(i, IndexCol) & "].bolSHT"
    Else: Cells(n, SHTDriverCol1) = "MDoor_A[" & Cells(i, IndexCol) & "].bolSHT"
    End If
    Cells(n, SHTDriverCol2) = FormatNumber(Cells(i, SHTCol), 2)
    'HCRDEnable
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(n, HCRDEnableDriverCol1) = "PDoor_A[" & Cells(i, IndexCol) & "].bolHCRDEnable"
        Cells(n, HCRDEnableDriverCol2) = FormatNumber(Cells(i, HCRDEnableCol), 2)
    End If
    'HUL
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(n, HULDriverCol1) = "PDoor_A[" & Cells(i, IndexCol) & "].bolHUL"
        Cells(n, HULDriverCol2) = FormatNumber(Cells(i, HULInputCol), 2)
    End If
    'SCRDEnable
    If Cells(i, LOKOutputCol) <> "" Then
        Cells(n, SCRDEnableDriverCol1) = "PDoor_A[" & Cells(i, IndexCol) & "].bolSCRDEnable"
        Cells(n, SCRDEnableDriverCol2) = FormatNumber(Cells(i, SCRDEnableCol), 2)
    End If
    n = n + 1
    'Output names and addresses for CX Programmer Symbols Page
    'EMR
    Cells(k, NamesCol) = "W" & Cells(i, DoorNameCol) & "_EMR" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "W" & FormatNumber(Cells(i, EMRCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "EMR" 'PLC bit Address
    k = k + 1
    'SHT
    Cells(k, NamesCol) = "W" & Cells(i, DoorNameCol) & "_SHT" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "W" & FormatNumber(Cells(i, SHTCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "SHT" 'PLC bit Address
    k = k + 1
    'HUL
    Cells(k, NamesCol) = "W" & Cells(i, DoorNameCol) & "_HUL" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "W" & FormatNumber(Cells(i, HULInputCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "HUL" 'PLC bit Address
    k = k + 1
    'OTL
    Cells(k, NamesCol) = "H" & Cells(i, DoorNameCol) & "_OTL" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "H" & FormatNumber(Cells(i, OTLOutputCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "OTL" 'PLC bit Address
    k = k + 1
    'BRC
    Cells(k, NamesCol) = "H" & Cells(i, DoorNameCol) & "_BRC" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "H" & FormatNumber(Cells(i, BRCOutputCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "BRC" 'PLC bit Address
    k = k + 1
    'CRD_H_IN
    Cells(k, NamesCol) = "W" & Cells(i, DoorNameCol) & "_CRD_H_IN" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "W" & FormatNumber(Cells(i, HCRDInputCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "CRD_H_IN" 'PLC bit Address
    k = k + 1
    'HCRDEnable
    Cells(k, NamesCol) = "W" & Cells(i, DoorNameCol) & "_HCRDEnable" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "W" & FormatNumber(Cells(i, HCRDEnableCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "HCRDEnable" 'PLC bit Address
    k = k + 1
    'CRD_S_IN
    Cells(k, NamesCol) = "W" & Cells(i, DoorNameCol) & "_CRD_S_IN" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "W" & FormatNumber(Cells(i, SCRDInputCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "CRD_S_IN" 'PLC bit Address
    k = k + 1
    'SCRDEnable
    Cells(k, NamesCol) = "W" & Cells(i, DoorNameCol) & "_SCRDEnable" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "W" & FormatNumber(Cells(i, SCRDEnableCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "SCRDEnable" 'PLC bit Address
    k = k + 1
    'REX_IN
    Cells(k, NamesCol) = "W" & Cells(i, DoorNameCol) & "_REX_IN" 'PLC bit name
    Cells(k, DataTypeCol) = "BOOL" 'PLC bit name
    Cells(k, AddressCol) = "W" & FormatNumber(Cells(i, ExitBtnCol), 2) 'PLC bit Address
    Cells(k, CommentCol) = "REX_IN" 'PLC bit Address
    k = k + 1
End If
Next
'Dim TargetRange As Range
'Dim cel As Range
'Set TargetRange = Range(Cells(1, lastCol + 5), Cells(300, 1000))
'TargetRange.Columns.AutoFit  'Fit the cells to the text
'HighestIndex = Enumerate(StartRow, IndexCol)
'For i = (lastCol + 5) To 300 '(lastCol + 5 + HighestIndex * ColumnOffset)
'    If Cells(2, i) = Empty Then Cells(2, i).EntireColumn.Hidden = True
'Next
End Sub

Public Sub IOListAutomation()
Dim i, NumOfObjects, StartRow, DeviceTypeCol, NameCol, InputOutputCol, SideCol, IndexCol, InAddressCol, OutAddressCol, CloseAddressCol, StopAddressCol As Integer
Dim AddressCol, CommentCol, ExternCommentCol, NamesCol, DataTypeCol, HighestIndex, DPSDriverCol, LOKDriverCol, sDPSDriverCol, sLOKDriverCol, PANDriverCol As Integer
Dim EMRDriverCol, OTLDriverCol, BRCDriverCol, SHTDriverCol, HCRDEnableDriverCol, HULDriverCol, SCRDEnableDriverCol, DeviceInfoCol, CLSDriverCol, STPDriverCol, OPNDriverCol, VHDDriverCol As Integer
Dim DoorMnemCol, CRDMnemCol, REXMnemCol, PANMnemCol, VHDMnemCol, KEYMnemCol, CPBMnemCol, ALMMnemCol, ELEVMnemCol, ERMMnemCol, FirstCol, MOTMnemCol, LastCol As Integer
Dim IndexDevice, iDevice, index, deviceType, Type_Index, Side As Variant
Dim mDPS, mCRD, mREX, mPAN, mVHD, mKEY, mCPB, mALM, mELEV, mERM, mMOT As Integer  'Mnemonics row
Dim iDPS, iCLS, iLOK, iHCRD, iSCRD, iREX, iPAN, iVHD, iKEY, iCPB, iALM, iELEV, iERM, iMOT, iSTP, iOPN, isDPS, isLOK As Integer  'Indusoft driver sheet row
Dim iBRC, iOTL, iEMR, iSHT, iHUL As Integer
Dim symbolRow As Integer
Dim TargetRange As Range

StartRow = 2    'The row just beneath the headers
IOStartRow = 5
'Offset values
Door_Open_Input = 0
DPS_Input = 25
Door_Open_Output = 50
Emergency_Release = 75
Shunt_Input = 100
Card_Reader_H_Input = 125
Card_Reader_H_Enable = 150
Hold_Unlock_Input = 175
Card_Reader_S_Input = 200
Card_Reader_S_Enable = 225
Exit_Push_Button = 250
Future = 275
Half_Cycle_Lock = 300
Future4 = 325
OTL_Output = 100
BRC_Output = 75
Door_Open_Input_Slider = 360
Door_Stop_Input_Slider = 361
Door_Close_Input_Slider = 362
Door_Shunt_Input_Slider = 363
Door_OTL_Output_Slider = 226.01
'Hold Bits for Vehicle, Key, Panic, etc
hVehicle = 280#
hKey = 290#
hPanic = 260#

'Initializing the mnemonic row position indexes
mDPS = StartRow
mCRD = StartRow
mREX = StartRow
mPAN = StartRow
mVHD = StartRow
mKEY = StartRow
mCPB = StartRow
mALM = StartRow
mELEV = StartRow
mERM = StartRow
mMOT = StartRow

'Initializing the Indusoft Driver row position indexes
iDPS = StartRow
iLOK = StartRow
isDPS = StartRow
isLOK = StartRow
iOPN = StartRow 'Slider open command (LOK for sliders)
iCLS = StartRow
iSTP = StartRow
iHCRD = StartRow
iSCRD = StartRow
iREX = StartRow
iPAN = StartRow
iVHD = StartRow
iKEY = StartRow
iCPB = StartRow
iALM = StartRow
iELEV = StartRow
iERM = StartRow
iMOT = StartRow
iBRC = StartRow
iOTL = StartRow
iEMR = StartRow
iSHT = StartRow
iHUL = StartRow

'Initializing the symbols row position indexes
symbolRow = StartRow

'Input Fields
DeviceTypeCol = 2   'Column number of the DoorNames Column
InputOutputCol = DeviceTypeCol + 1
NameCol = InputOutputCol + 1    'IndusoftIndex Column
SideCol = NameCol + 1    'PLC address of the EMR bit Column - Door_Open_Input Ex: W2D209_EMR 75.00
IndexCol = SideCol + 1
InAddressCol = IndexCol + 1
OutAddressCol = InAddressCol + 1
StopAddressCol = OutAddressCol + 1
CloseAddressCol = StopAddressCol + 1
DeviceInfoCol = CloseAddressCol + 1

FirstCol = InAddressCol + 3

'Output Fields
'Symbols columns
ExternCommentCol = InAddressCol + 9
NamesCol = ExternCommentCol + 1
DataTypeCol = NamesCol + 1
AddressCol = DataTypeCol + 1
CommentCol = AddressCol + 1
'Mnemonics columns
DoorMnemCol = CommentCol + 2  'Associated last row index: mDPS
CRDMnemCol = DoorMnemCol + 2    'mCRD
REXMnemCol = CRDMnemCol + 2     'mREX
PANMnemCol = REXMnemCol + 2     'mPAN
VHDMnemCol = PANMnemCol + 2     'mVHD
KEYMnemCol = VHDMnemCol + 2     'mKEY
CPBMnemCol = KEYMnemCol + 2     'mCPB
ALMMnemCol = CPBMnemCol + 2     'mALM
ELEVMnemCol = ALMMnemCol + 2    'mELE
ERMMnemCol = ELEVMnemCol + 2    'mERM
MOTMnemCol = ERMMnemCol + 2    'mMOT
'Indusoft Driver columns
DPSDriverCol = MOTMnemCol + 2
LOKDriverCol = DPSDriverCol + 3
sDPSDriverCol = LOKDriverCol + 3
sLOKDriverCol = sDPSDriverCol + 3
OPNDriverCol = sLOKDriverCol + 3
STPDriverCol = OPNDriverCol + 3
CLSDriverCol = STPDriverCol + 3
BRCDriverCol = CLSDriverCol + 3
OTLDriverCol = BRCDriverCol + 3
EMRDriverCol = OTLDriverCol + 3
SHTDriverCol = EMRDriverCol + 3
HULDriverCol = SHTDriverCol + 3
HCRDEnableDriverCol = HULDriverCol + 3
SCRDEnableDriverCol = HCRDEnableDriverCol + 3
VHDDriverCol = SCRDEnableDriverCol + 3
PANDriverCol = VHDDriverCol + 3
LastCol = PANDriverCol + 3

'Write headers for each output column
Cells(1, ExternCommentCol) = "External Comment (Do not use)"
Cells(1, NamesCol) = "Names"
Cells(1, DataTypeCol) = "Data Type"
Cells(1, AddressCol) = "Address"
Cells(1, CommentCol) = "Comment"
Cells(1, DoorMnemCol) = "Door Mnemonics (DPS)"
Cells(1, CRDMnemCol) = "Card Reader mnemonics (CRD)"
Cells(1, REXMnemCol) = "Request Exit Mnemonics (REX)"
Cells(1, PANMnemCol) = "Panic/Duress Mnemonics (PAN)"
Cells(1, VHDMnemCol) = "Vehicle Detector Mnemonics (VHD)"
Cells(1, KEYMnemCol) = "Key Mnemonics (KEY)"
Cells(1, CPBMnemCol) = "Call Pushbutton (CPB)"
Cells(1, ALMMnemCol) = "Alarm mnemonics (ALM)"
Cells(1, ELEVMnemCol) = "Elevator Mnemonics (ELEV)"
Cells(1, ERMMnemCol) = "Alarm Sounder/Strobe Mnemonics (ERM)"
Cells(1, MOTMnemCol) = "Motion Detector mnemonics (MOT)"
Cells(1, DPSDriverCol) = "Indusoft Driver Sheet - DPS (W)"
Cells(1, LOKDriverCol) = "Indusoft Driver Sheet - LOK (W)"
Cells(1, sDPSDriverCol) = "Indusoft Driver Sheet - Slider DPS (W)"
Cells(1, sLOKDriverCol) = "Indusoft Driver Sheet - Slider LOK (W)"
Cells(1, OPNDriverCol) = "Indusoft Driver Sheet - Slider Open (W)"
Cells(1, STPDriverCol) = "Indusoft Driver Sheet - STP (W)"
Cells(1, CLSDriverCol) = "Indusoft Driver Sheet - CLS (W)"
Cells(1, BRCDriverCol) = "Indusoft Driver BRC (H)"
Cells(1, OTLDriverCol) = "Indusoft Driver OTL (H)"
Cells(1, EMRDriverCol) = "Indusoft Driver EMR (W)"
Cells(1, SHTDriverCol) = "Indusoft Driver SHT (W)"
Cells(1, HULDriverCol) = "Indusoft Driver HUL (W)"
Cells(1, HCRDEnableDriverCol) = "Indusoft Driver HCRDEnable (W)"
Cells(1, SCRDEnableDriverCol) = "Indusoft Driver SCRDEnable (W)"
Cells(1, VHDDriverCol) = "Indusoft Driver Vehicle Detector (H)"
Cells(1, PANDriverCol) = "Indusoft Driver Panic (H)"

NumOfObjects = Enumerate(StartRow, NameCol)  'Number of of Objects pasted in by the user.

For i = StartRow To NumOfObjects    'Cycling through the IO items
    deviceType = Cells(i, DeviceTypeCol)    'DPS, or CRD, or REX
    Side = Cells(i, SideCol)    'H or S
    Type_Index = Cells(i, IndexCol)
    If Len(Type_Index) > 1 And Type_Index <> "N/A" Then
        IndexDevice = Split(Type_Index, " ")  'split "p 1"
        iDevice = IndexDevice(0)              ' into "p"
        index = IndexDevice(1)                'and "1"
    Else
        iDevice = ""
        index = ""
    End If
    Select Case Cells(i, DeviceTypeCol)
        Case "DPS":
            If iDevice = "p" Then    'PDoors
                'Do mnemonics
                Call Mnemonics(i, deviceType, iDevice, mDPS, DoorMnemCol, Side, index)  'Current Device Row, 'DPS', 'p', row, column, H or S side, Indusoft Index
                mDPS = mDPS + 5 'A PDoor mnemonic takes up 5 lines
                
                'Do symbols sheet
                Call SymbolSheet(i, deviceType, iDevice, symbolRow, NamesCol, Side)
                'symbolRow = symbolRow + 9 'There are 4 PDoor-related symbols
                
                'Do Indusoft Driver sheets for DPS and LOK
                Call DriverSheet(i, "DPS", iDevice, iDPS, DPSDriverCol, index, Side)
                iDPS = iDPS + 1
                Call DriverSheet(i, "LOK", iDevice, iLOK, LOKDriverCol, index, Side)
                iLOK = iLOK + 1
                Call DriverSheet(i, "BRC", iDevice, iBRC, BRCDriverCol, index, Side)
                iBRC = iBRC + 1
                Call DriverSheet(i, "OTL", iDevice, iOTL, OTLDriverCol, index, Side)
                iOTL = iOTL + 1
                Call DriverSheet(i, "EMR", iDevice, iEMR, EMRDriverCol, index, Side)
                iEMR = iEMR + 1
                Call DriverSheet(i, "SHT", iDevice, iSHT, SHTDriverCol, index, Side)
                iSHT = iSHT + 1
                Call DriverSheet(i, "HUL", iDevice, iHUL, HULDriverCol, index, Side)
                iHUL = iHUL + 1
                
            ElseIf iDevice = "m" Then     'MDoors
                Call Mnemonics(i, deviceType, iDevice, mDPS, DoorMnemCol, Side, index)
                mDPS = mDPS + 3
                Call DriverSheet(i, "BRC", iDevice, iBRC, BRCDriverCol, index, Side)
                iBRC = iBRC + 1
                Call DriverSheet(i, "OTL", iDevice, iOTL, OTLDriverCol, index, Side)
                iOTL = iOTL + 1
                Call DriverSheet(i, "EMR", iDevice, iEMR, EMRDriverCol, index, Side)
                iEMR = iEMR + 1
                Call DriverSheet(i, "SHT", iDevice, iSHT, SHTDriverCol, index, Side)
                iSHT = iSHT + 1
                Call SymbolSheet(i, deviceType, iDevice, symbolRow, NamesCol, Side)
                'symbolRow = symbolRow + 5 'There are 2 MDoor-related symbols
                Call DriverSheet(i, "DPS", iDevice, iDPS, DPSDriverCol, index, Side)
                iDPS = iDPS + 1
            ElseIf iDevice = "o" Then     'OSCDoors
                'Call Mnemonics(i, deviceType, iDevice, mDPS, DoorMnemCol, Side, Index)
                'mDPS = mDPS + 1
                Call DriverSheet(i, "sDPS", iDevice, isDPS, sDPSDriverCol, index, Side)
                isDPS = isDPS + 1
                Call DriverSheet(i, "sLOK", iDevice, isLOK, sLOKDriverCol, index, Side)
                isLOK = isLOK + 1
                Call DriverSheet(i, "OPN", iDevice, iOPN, OPNDriverCol, index, Side)
                iOPN = iOPN + 1
                Call DriverSheet(i, "CLS", iDevice, iCLS, CLSDriverCol, index, Side)
                iCLS = iCLS + 1
                If Cells(i, StopAddressCol) <> "" Then
                    Call DriverSheet(i, "STP", iDevice, iSTP, STPDriverCol, index, Side)
                    iSTP = iSTP + 1
                End If
                'Call DriverSheet(i, "sBRC", iDevice, isBRC, sBRCDriverCol, Index, Side)
                'isBRC = isBRC + 1
                'Call DriverSheet(i, "sOTL", iDevice, isOTL, sOTLDriverCol, Index, Side)
                'isOTL = isOTL + 1
                'Call DriverSheet(i, "sSHT", iDevice, isSHT, sSHTDriverCol, Index, Side)
                'isSHT = isSHT + 1
                'Call DriverSheet(i, "sHUL", iDevice, isHUL, sHULDriverCol, Index, Side)
                'isHUL = isHUL + 1

                'isHUL = isHUL + 1
            End If
        Case "CRD":
            Call Mnemonics(i, deviceType, iDevice, mCRD, CRDMnemCol, Side, index)
            mCRD = mCRD + 4
            Call SymbolSheet(i, deviceType, iDevice, symbolRow, NamesCol, Side)
            If Side = "H" Then
                Call DriverSheet(i, "CRD", iDevice, iHCRD, HCRDEnableDriverCol, index, Side)
                iHCRD = iHCRD + 1
            ElseIf Side = "S" Then
                Call DriverSheet(i, "CRD", iDevice, iSCRD, SCRDEnableDriverCol, index, Side)
                iSCRD = iSCRD + 1
            End If
        Case "REX":
            Call Mnemonics(i, deviceType, iDevice, mREX, REXMnemCol, Side, index)
            mREX = mREX + 3
        Case "KEY":
            Call Mnemonics(i, deviceType, iDevice, mKEY, KEYMnemCol, Side, index)
            mKEY = mKEY + 3
        Case "PAN":
            Call Mnemonics(i, deviceType, iDevice, mPAN, PANMnemCol, Side, index)
            mPAN = mPAN + 3
            Call DriverSheet(i, "PAN", iDevice, iPAN, PANDriverCol, index, Side)
            iPAN = iPAN + 1
            hPanic = hPanic + 0.01
        Case "VHD":
            Call Mnemonics(i, deviceType, iDevice, mVHD, VHDMnemCol, Side, index)
            mVHD = mVHD + 3
            Call DriverSheet(i, "VHD", iDevice, iVHD, VHDDriverCol, index, Side)
            iVHD = iVHD + 1
            hVehicle = hVehicle + 0.01
        Case "CPB":
            Call Mnemonics(i, deviceType, iDevice, mCPB, CPBMnemCol, Side, index)
            mCPB = mCPB + 1
        Case "ALM":
            Call Mnemonics(i, deviceType, iDevice, mALM, ALMMnemCol, Side, index)
            mALM = mALM + 3
        Case "ELEV":
            Call Mnemonics(i, deviceType, iDevice, mELEV, ELEVMnemCol, Side, index)
            mELEV = mELEV + 1
        Case "ERM":
            Call Mnemonics(i, deviceType, iDevice, mERM, ERMMnemCol, Side, index)
            mERM = mERM + 1
        Case "MOT":
            Call Mnemonics(i, deviceType, iDevice, mMOT, MOTMnemCol, Side, index)
            mMOT = mMOT + 3
    End Select
Next
Set TargetRange = Range(Cells(1, FirstCol), Cells(mDPS, LastCol))
TargetRange.Columns.AutoFit  'Fit the cells to the text
End Sub

Public Sub Mnemonics(objectRow, deviceType, iDevice, row, Col, Side, index)
Dim k, v, DeviceTypeCol, InputOutputCol, NameCol, SideCol, IndexCol, InAddressCol, OutAddressCol, CloseAddressCol, StopAddressCol, DeviceInfoCol As Integer

'Input Fields
StartRow = 2
DeviceTypeCol = 2   'Column number of the DoorNames Column
InputOutputCol = DeviceTypeCol + 1
NameCol = InputOutputCol + 1    'IndusoftIndex Column
SideCol = NameCol + 1    'PLC address of the EMR bit Column - Door_Open_Input Ex: W2D209_EMR 75.00
IndexCol = SideCol + 1
InAddressCol = IndexCol + 1
OutAddressCol = InAddressCol + 1
StopAddressCol = OutAddressCol + 1
CloseAddressCol = StopAddressCol + 1
DeviceInfoCol = CloseAddressCol + 1

Select Case deviceType
    Case "DPS":
        Select Case iDevice
            Case "p":
                Cells(row, Col) = " ' PDoor - " & Cells(objectRow, NameCol) & ", Indusoft Index: " & iDevice & index 'Rung comment section
                Cells(row + 1, Col) = "LD " & FormatNumber(Cells(objectRow, InAddressCol), 2)
                Cells(row + 2, Col) = "OUT W" & FormatNumber(Cells(objectRow, InAddressCol) + DPS_Input, 2)
                Cells(row + 3, Col) = "LD W" & FormatNumber(Cells(objectRow, InAddressCol) + Door_Open_Output, 2)
                Cells(row + 4, Col) = "OUT " & FormatNumber(Cells(objectRow, OutAddressCol), 2)
            Case "m":
                Cells(row, Col) = " ' MDoor - " & Cells(objectRow, NameCol) & ", Indusoft Index: " & iDevice & index  'Rung comment section
                Cells(row + 1, Col) = "LD " & FormatNumber(Cells(objectRow, InAddressCol), 2)
                Cells(row + 2, Col) = "OUT W" & FormatNumber(Cells(objectRow, InAddressCol) + DPS_Input, 2)
            Case "o":
                'Cells(row, Col) = deviceType & " " & iDevice
        End Select
    Case "CRD":
        Select Case iDevice
            Case "p":
                Select Case Side
                    Case "H":
                        Cells(row, Col) = " ' H-side Card Reader - " & Cells(objectRow, NameCol) & ", Indusoft Index: " & iDevice & index 'Rung comment section"
                        Cells(row + 1, Col) = "LD " & FormatNumber(Cells(objectRow, InAddressCol), 2)
                        For k = StartRow To InputNumOfObjects   'hunt down the door this card reader belongs to to get its input address
                            If Cells(k, DeviceTypeCol) = "DPS" And Cells(k, NameCol) = Cells(objectRow, NameCol) Then
                                Cells(row + 2, Col) = "AND W" & Cells(k, InAddressCol) + Card_Reader_H_Enable
                                Cells(row + 3, Col) = "OUT W" & Cells(k, InAddressCol) + Card_Reader_H_Input
                                k = InputNumOfObjects + 1
                            End If
                        Next
                    Case "S":
                        Cells(row, Col) = " ' S-side Card Reader - " & Cells(objectRow, NameCol) & ", Indusoft Index: " & iDevice & index 'Rung comment section"
                        Cells(row + 1, Col) = "LD " & FormatNumber(Cells(objectRow, InAddressCol), 2)
                        For k = StartRow To InputNumOfObjects
                            If Cells(k, DeviceTypeCol) = "DPS" And Cells(k, NameCol) = Cells(objectRow, NameCol) Then
                                Cells(row + 2, Col) = "AND W" & Cells(k, InAddressCol) + Card_Reader_S_Enable
                                Cells(row + 3, Col) = "OUT W" & Cells(k, InAddressCol) + Card_Reader_S_Input
                                k = InputNumOfObjects + 1
                            End If
                        Next
                End Select
            Case "o":
        End Select
    Case "REX":
        Select Case iDevice
            Case "p", "o":
                Cells(row, Col) = " ' Request to exit - " & Cells(objectRow, NameCol) & ", Indusoft Index: " & iDevice & index 'Rung comment section"
                Cells(row + 1, Col) = "LD " & FormatNumber(Cells(objectRow, InAddressCol), 2)
                'Find the door whom this REX belongs to to get its input address
                For k = StartRow To InputNumOfObjects
                    If Cells(k, DeviceTypeCol) = "DPS" And Cells(k, NameCol) = Cells(objectRow, NameCol) Then
                        Cells(row + 2, Col) = "OUT W" & FormatNumber(Cells(k, InAddressCol) + Exit_Push_Button, 2)
                    End If
                Next
        End Select
    Case "KEY":
        Select Case iDevice
            Case "p", "o":
                Cells(row, Col) = " ' Keyed door - " & Cells(objectRow, NameCol) & ", Indusoft Index: " & iDevice & index 'Rung comment section"
                Cells(row + 1, Col) = "LD " & FormatNumber(Cells(objectRow, InAddressCol), 2)
                'Find the door whom this REX belongs to to get its input address
                For k = StartRow To InputNumOfObjects
                    If Cells(k, DeviceTypeCol) = "DPS" And Cells(k, NameCol) = Cells(objectRow, NameCol) Then
                        Cells(row + 2, Col) = "OUT W" & FormatNumber(Cells(k, InAddressCol) + Exit_Push_Button, 2)  'REX has the basic functionality, so they share a memory space
                    End If
                Next
        End Select
    Case "PAN":
        Cells(row, Col) = " ' Panic - " & Cells(objectRow, DeviceInfoCol) & ", Indusoft Index: " & iDevice & index 'Rung comment section"
        Cells(row + 1, Col) = "LDNOT " & FormatNumber(Cells(objectRow, InAddressCol), 2)
        Cells(row + 2, Col) = "SET H" & hPanic
    Case "VHD":
        Cells(row, Col) = " ' Vehicle detector - " & Cells(objectRow, DeviceInfoCol) & ", Indusoft Index: " & iDevice & index 'Rung comment section"
        Cells(row + 1, Col) = "LD " & FormatNumber(Cells(objectRow, InAddressCol), 2)
        Cells(row + 2, Col) = "SET H" & hVehicle
    Case "CPB":
    Case "ALM":
        Cells(row, Col) = " ' Alarm - " & Cells(objectRow, DeviceInfoCol) & ", Indusoft Index: " & iDevice & index 'Rung comment section"
        Cells(row + 1, Col) = "LD " & FormatNumber(Cells(objectRow, InAddressCol), 2)
        Cells(row + 2, Col) = "SET H" & FormatNumber(Cells(objectRow, InAddressCol) + 250, 2)
    Case "ELEV":
    Case "ERM":
    Case "MOT":
        Select Case iDevice
            Case "p", "o":
                Cells(row, Col) = " ' Motion Request to exit - " & Cells(objectRow, NameCol) & ", Indusoft Index: " & iDevice & index 'Rung comment section"
                Cells(row + 1, Col) = "LD " & FormatNumber(Cells(objectRow, InAddressCol), 2)
                'Find the door whom this MOT belongs to to get its input address
                For k = StartRow To InputNumOfObjects
                    If Cells(k, DeviceTypeCol) = "DPS" And Cells(k, NameCol) = Cells(objectRow, NameCol) Then
                        Cells(row + 2, Col) = "OUT W" & FormatNumber(Cells(k, InAddressCol) + Exit_Push_Button, 2)
                    End If
                Next
        End Select
End Select
End Sub

Public Sub SymbolSheet(objectRow, deviceType, iDevice, row, Col, Side)
Dim DeviceTypeCol, InputOutputCol, NameCol, SideCol, IndexCol, InAddressCol, OutAddressCol, CloseAddressCol, StopAddressCol, DeviceInfoCol As Integer
Dim index As String
'Input Fields
DeviceTypeCol = 2   'Column number of the DoorNames Column
InputOutputCol = DeviceTypeCol + 1
NameCol = InputOutputCol + 1    'IndusoftIndex Column
SideCol = NameCol + 1    'PLC address of the EMR bit Column - Door_Open_Input Ex: W2D209_EMR 75.00
IndexCol = SideCol + 1
InAddressCol = IndexCol + 1
OutAddressCol = InAddressCol + 1
StopAddressCol = OutAddressCol + 1
CloseAddressCol = StopAddressCol + 1
DeviceInfoCol = CloseAddressCol + 1

Cells(row, Col - 1) = deviceType & " " & Cells(objectRow, IndexCol) & " - Name: " & Cells(objectRow, NameCol)    'External comment column
Select Case deviceType
    Case "DPS":
        Select Case iDevice
            Case "p":
                Cells(row, Col) = "I" & Cells(objectRow, NameCol) & "_DPS"              'Name
                Cells(row, Col + 1) = "BOOL"                                            'Datatype
                Cells(row, Col + 2) = FormatNumber(Cells(objectRow, InAddressCol), 2)   'Address
                Cells(row, Col + 3) = "DPS Input"                                       'Comment
                row = row + 1
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_DPS"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + DPS_Input, 2)
                Cells(row, Col + 3) = "DPS"
                row = row + 1
                Cells(row, Col) = "WO" & Cells(objectRow, NameCol) & "_LOK"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + Door_Open_Output, 2)
                Cells(row, Col + 3) = "LOK"
                row = row + 1
                Cells(row, Col) = "O" & Cells(objectRow, NameCol) & "_LOK"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = FormatNumber(Cells(objectRow, OutAddressCol), 2)
                Cells(row, Col + 3) = "LOK Output"
                row = row + 1
                'EMR
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_EMR"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + Emergency_Release, 2)
                Cells(row, Col + 3) = "EMR"
                row = row + 1
                'SHT
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_SHT"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + Shunt_Input, 2)
                Cells(row, Col + 3) = "SHT"
                row = row + 1
                'HUL
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_HUL"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + Hold_Unlock_Input, 2)
                Cells(row, Col + 3) = "HUL"
                row = row + 1
                'OTL
                Cells(row, Col) = "H" & Cells(objectRow, NameCol) & "_OTL"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "H" & FormatNumber(Cells(objectRow, InAddressCol) + OTL_Output, 2)
                Cells(row, Col + 3) = "OTL"
                row = row + 1
                'BRC
                Cells(row, Col) = "H" & Cells(objectRow, NameCol) & "_BRC"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "H" & FormatNumber(Cells(objectRow, InAddressCol) + BRC_Output, 2)
                Cells(row, Col + 3) = "BRC"
                row = row + 1
                'CRD_H_IN
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_CRD_H_IN"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + Card_Reader_H_Input, 2)
                Cells(row, Col + 3) = "CRD_H_IN"
                row = row + 1
                'HCRDEnable
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_HCRDEnable"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + Card_Reader_H_Enable, 2)
                Cells(row, Col + 3) = "HCRDEnable"
                row = row + 1
                'CRD_S_IN
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_CRD_S_IN"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + Card_Reader_S_Input, 2)
                Cells(row, Col + 3) = "CRD_S_IN"
                row = row + 1
                'SCRDEnable
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_SCRDEnable"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + Card_Reader_S_Enable, 2)
                Cells(row, Col + 3) = "SCRDEnable"
                row = row + 1
                'REX_IN
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_REX_IN"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + Exit_Push_Button, 2)
                Cells(row, Col + 3) = "REX_IN"
                row = row + 1
            Case "m":
                Cells(row, Col) = "I" & Cells(objectRow, NameCol) & "_DPS"              'Name
                Cells(row, Col + 1) = "BOOL"                                            'Datatype
                Cells(row, Col + 2) = FormatNumber(Cells(objectRow, InAddressCol), 2)   'Address
                Cells(row, Col + 3) = "DPS Input"                                       'Comment
                row = row + 1
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_DPS"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + DPS_Input, 2)
                Cells(row, Col + 3) = "DPS"
                row = row + 1
                'SHT
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_SHT"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + Shunt_Input, 2)
                Cells(row, Col + 3) = "SHT"
                row = row + 1
                'OTL
                Cells(row, Col) = "H" & Cells(objectRow, NameCol) & "_OTL"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "H" & FormatNumber(Cells(objectRow, InAddressCol) + OTL_Output, 2)
                Cells(row, Col + 3) = "OTL"
                row = row + 1
                'BRC
                Cells(row, Col) = "H" & Cells(objectRow, NameCol) & "_BRC"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "H" & FormatNumber(Cells(objectRow, InAddressCol) + BRC_Output, 2)
                Cells(row, Col + 3) = "BRC"
            Case "o":
                Cells(row, Col) = deviceType & " " & iDevice
        End Select
    Case "CRD":
        Select Case Side
            Case "H":
                'CRD_H_IN
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_CRD_H_IN"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol), 2)
                Cells(row, Col + 3) = "CRD_H_IN"
                'HCRDEnable
                Cells(row + 1, Col) = "W" & Cells(objectRow, NameCol) & "_HCRDEnable"
                Cells(row + 1, Col + 1) = "BOOL"
                Cells(row + 1, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + 25, 2)
                Cells(row + 1, Col + 3) = "HCRDEnable"
            Case "S":
                'CRD_S_IN
                Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_CRD_S_IN"
                Cells(row, Col + 1) = "BOOL"
                Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol), 2)
                Cells(row, Col + 3) = "CRD_S_IN"
                'SCRDEnable
                Cells(row + 1, Col) = "W" & Cells(objectRow, NameCol) & "_SCRDEnable"
                Cells(row + 1, Col + 1) = "BOOL"
                Cells(row + 1, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol) + 25, 2)
                Cells(row + 1, Col + 3) = "SCRDEnable"
        End Select
    Case "REX":
        'REX_IN
        Cells(row, Col) = "W" & Cells(objectRow, NameCol) & "_REX_IN"
        Cells(row, Col + 1) = "BOOL"
        Cells(row, Col + 2) = "W" & FormatNumber(Cells(objectRow, InAddressCol), 2)
        Cells(row, Col + 3) = "REX_IN"
    Case "PAN":
    Case "VHD":
    Case "KEY":
    Case "CPB":
    Case "ALM":
    Case "ELEV":
    Case "ERM":
    Case "MOT":
End Select
End Sub

Public Sub DriverSheet(objectRow, deviceType, iDevice, row, Col, ind, Side)
Dim k, DeviceTypeCol, InputOutputCol, NameCol, SideCol, IndexCol, InAddressCol, OutAddressCol, CloseAddressCol, StopAddressCol As Integer
Dim index As String
'Input Fields
DeviceTypeCol = 2   'Column number of the DoorNames Column
InputOutputCol = DeviceTypeCol + 1
NameCol = InputOutputCol + 1    'IndusoftIndex Column
SideCol = NameCol + 1    'PLC address of the EMR bit Column - Door_Open_Input Ex: W2D209_EMR 75.00
IndexCol = SideCol + 1
InAddressCol = IndexCol + 1
OutAddressCol = InAddressCol + 1
StopAddressCol = OutAddressCol + 1
CloseAddressCol = StopAddressCol + 1

Select Case deviceType
    Case "DPS":
        Select Case iDevice
            Case "p":
                Cells(row, Col) = "PDoor[" & ind & "].bolDPS"
                'Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + 25, 2)
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + DPS_Input, 2)
            Case "m":
                Cells(row, Col) = "MDoor[" & ind & "].bolDPS"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + DPS_Input, 2)
        End Select
    Case "sDPS":
        Select Case iDevice
            Case "o":
                Cells(row, Col) = "OSCDoor[" & ind & "].bolDPS"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + DPS_Input, 2)
            Case "v":
                Cells(row, Col) = "OSCDoor[" & ind & "].bolDPS"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + DPS_Input, 2)
        End Select
    Case "OPN":
        Cells(row, Col) = "OSCDoor[" & ind & "].bolLOK"
        Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol), 2)
    Case "CLS":
        Cells(row, Col) = "OSCDoor[" & ind & "].bolCLS"
        Cells(row, Col + 1) = FormatNumber(Cells(objectRow, CloseAddressCol), 2)
    Case "STP":
        Cells(row, Col) = "OSCDoor[" & ind & "].bolSTP"
        Cells(row, Col + 1) = FormatNumber(Cells(objectRow, StopAddressCol), 2)
    Case "LOK":
        Select Case iDevice
            Case "p":
                Cells(row, Col) = "PDoor[" & ind & "].bolLOK"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol), 2)
        End Select
    Case "sLOK":
        Select Case iDevice
            Case "o", "v":
                Cells(row, Col) = "OSCDoor[" & ind & "].bolLOK"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + Door_Open_Input_Slider, 2)
        End Select
    Case "CRD":
        Select Case iDevice
        Case "p":
            Select Case Side
                Case "H":
                    Cells(row, Col) = "PDoor[" & ind & "].bolHCRDEnable"
                    For k = StartRow To InputNumOfObjects
                        If Cells(k, DeviceTypeCol) = "DPS" And Cells(k, NameCol) = Cells(objectRow, NameCol) Then
                            Cells(row, Col + 1) = FormatNumber(Cells(k, InAddressCol) + Card_Reader_H_Enable, 2)
                        End If
                    Next
                Case "S":
                    Cells(row, Col) = "PDoor[" & ind & "].bolSCRDEnable"
                    For k = StartRow To InputNumOfObjects
                        If Cells(k, DeviceTypeCol) = "DPS" And Cells(k, NameCol) = Cells(objectRow, NameCol) Then
                            Cells(row, Col + 1) = FormatNumber(Cells(k, InAddressCol) + Card_Reader_S_Enable, 2)
                        End If
                    Next
            End Select
        Case "o":
        End Select
    Case "REX":
        Select Case iDevice
            Case "p":
                Cells(row, Col) = deviceType & " " & iDevice
        End Select
    Case "PAN":
            Cells(row, Col) = "Panic[" & ind & "].intPanic"
            Cells(row, Col + 1) = FormatNumber(hPanic, 2)
            'Cells(row, Col + 1) = "Fill in"
    Case "VHD":
            Cells(row, Col) = "Vehicle[" & ind & "].bolVHD"
            Cells(row, Col + 1) = FormatNumber(hVehicle, 2)
            'Cells(row, Col + 1) = "Fill in"
    Case "KEY":
    Case "CPB":
    Case "ALM":
    Case "ELEV":
    Case "ERM":
    Case "MOT":
    Case "BRC":
        Select Case iDevice
            Case "p":
                Cells(row, Col) = "PDoor[" & ind & "].bolBRC"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + BRC_Output, 2)
            Case "m":
                Cells(row, Col) = "MDoor[" & ind & "].bolBRC"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + BRC_Output, 2)
            Case "o":
                Cells(row, Col) = "OSCDoor[" & ind & "].bolBRC"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + BRC_Output, 2)
        End Select
    Case "OTL":
        Select Case iDevice
            Case "p":
                Cells(row, Col) = "PDoor[" & ind & "].bolOTL"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + OTL_Output, 2)
            Case "m":
                Cells(row, Col) = "MDoor[" & ind & "].bolOTL"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + OTL_Output, 2)
            Case "o":
                Cells(row, Col) = "OSCDoor[" & ind & "].bolOTL"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + Door_OTL_Output_Slider, 2)
        End Select
    Case "EMR":
        Select Case iDevice
            Case "p":
                Cells(row, Col) = "PDoor[" & ind & "].bolEMR"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + Emergency_Release, 2)
            Case "o":
                Cells(row, Col) = "OSCDoor[" & ind & "].bolEMR"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + Emergency_Release, 2)
        End Select
    Case "SHT":
        Select Case iDevice
            Case "p":
                Cells(row, Col) = "PDoor[" & ind & "].bolSHT"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + Shunt_Input, 2)
            Case "m":
                Cells(row, Col) = "MDoor[" & ind & "].bolSHT"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + Shunt_Input, 2)
            Case "o":
                Cells(row, Col) = "OSCDoor[" & ind & "].bolSHT"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + Door_Shunt_Input_Slider, 2)
        End Select
    Case "HUL":
        Select Case iDevice
            Case "p":
                Cells(row, Col) = "PDoor[" & ind & "].bolHUL"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + Hold_Unlock_Input, 2)
            Case "o":
                Cells(row, Col) = "OSCDoor[" & ind & "].bolHUL"
                Cells(row, Col + 1) = FormatNumber(Cells(objectRow, InAddressCol) + Hold_Unlock_Input, 2)
        End Select
End Select
End Sub

Public Sub IOListImport()
Dim Autolist As Workbook
Dim IOList As Workbook
Dim IOListPath As String
IOListPath = "IOList.xlsx"
Dim i, j, k, LocalStartRow, LocalInputDeviceCol, InDevice, InName, OutName, InputDeviceCol, OutputDevice, OutputDeviceCol As Integer
Dim InputNameCol, OutputNameCol, LocalNameCol, InputSideCol, OutputSideCol, LocalSideCol, IOIndexCol, LocalIndexCol As Integer
Dim InputAddressCol, LocalInputOutput, OutputAddressCol, LocalInputAddressCol, LocalOutputAddressCol, LocalCloseAddressCol, LocalStopAddressCol As Integer
Dim InputDeviceInfo, OutputDeviceInfo, LocalDeviceInfo As Integer
Dim index, Match As Integer
Dim SheetName, Type_Index, IndexDevice, iDevice, Name, Device, OutDevice As String
Dim TargetRange As Range
SheetName = Cells(1, 1)
LocalStartRow = 2
IOStartRow = 5
InputDeviceCol = 1
OutputDeviceCol = 17
InputNameCol = 3
OutputNameCol = 19
InputSideCol = 4
OutputSideCol = 20
IOIndexCol = 15
InputAddressCol = 9
OutputAddressCol = 25
InputDeviceInfo = 2
OutputDeviceInfo = 18

'Local Column placement
LocalInputDeviceCol = 2
LocalInputOutput = LocalInputDeviceCol + 1
LocalNameCol = LocalInputOutput + 1
LocalSideCol = LocalNameCol + 1
LocalIndexCol = LocalSideCol + 1
LocalInputAddressCol = LocalIndexCol + 1
LocalOutputAddressCol = LocalInputAddressCol + 1
LocalStopAddressCol = LocalOutputAddressCol + 1
LocalCloseAddressCol = LocalStopAddressCol + 1
LocalDeviceInfo = LocalCloseAddressCol + 1


'Put down column headers
Cells(1, LocalInputDeviceCol) = "Device"
Cells(1, LocalInputOutput) = "Input/Output"
Cells(1, LocalNameCol) = "Name"
Cells(1, LocalSideCol) = "Side"
Cells(1, LocalIndexCol) = "Indusoft Index"
Cells(1, LocalInputAddressCol) = "Input Address"
Cells(1, LocalOutputAddressCol) = "Output Address"
Cells(1, LocalCloseAddressCol) = "Close Address"
Cells(1, LocalStopAddressCol) = "Stop Address"
Cells(1, LocalDeviceInfo) = "Device Info"

InputNumOfObjects = OtherSheetEnumerate(SheetName, IOStartRow, InputDeviceCol)
OutputNumOfObjects = OtherSheetEnumerate(SheetName, IOStartRow, OutputDeviceCol)
j = LocalStartRow
k = LocalStartRow
'Import Inputs
For i = IOStartRow To InputNumOfObjects + IOStartRow - 1
    If Sheets(SheetName).Cells(i, InputDeviceCol) <> "SPR" Then
        ActiveSheet.Cells(j, LocalInputDeviceCol) = Sheets(SheetName).Cells(i, InputDeviceCol)  'Fill in the "DPS" cell
        ActiveSheet.Cells(j, LocalInputOutput) = "Input"  'Fill in the Input/Output cell
        ActiveSheet.Cells(j, LocalNameCol) = Sheets(SheetName).Cells(i, InputNameCol)   'Fill in the device name cell
        ActiveSheet.Cells(j, LocalSideCol) = Sheets(SheetName).Cells(i, InputSideCol)   'fill in the H/S side cell
        ActiveSheet.Cells(j, LocalIndexCol) = Sheets(SheetName).Cells(i, IOIndexCol)    'Fill in the "p 1" Indusoft Index cell
        ActiveSheet.Cells(j, LocalInputAddressCol) = Sheets(SheetName).Cells(i, InputAddressCol)    'Fill in the input address cell
        ActiveSheet.Cells(j, LocalDeviceInfo) = Sheets(SheetName).Cells(i, InputDeviceInfo)    'Fill in the device info cell
        'Do the door outputs
        Type_Index = Cells(j, LocalIndexCol)
        Device = Cells(j, LocalInputDeviceCol)
        Name = Cells(j, LocalNameCol)
        If Len(Type_Index) > 1 Then
            IndexDevice = Split(Type_Index, " ")  'split "p 1"
            iDevice = IndexDevice(0)              ' into "p"
            index = IndexDevice(1)                ' and 1
            If Device = "DPS" Then
                For k = IOStartRow To OutputNumOfObjects
                    OutDevice = Sheets(SheetName).Cells(k, OutputDeviceCol)
                    Select Case iDevice
                        Case "p":
                            If (Sheets(SheetName).Cells(k, OutputNameCol) = Name) And (OutDevice = "LOK") Then
                                ActiveSheet.Cells(j, LocalOutputAddressCol) = Sheets(SheetName).Cells(k, OutputAddressCol)
                                k = OutputNumOfObjects + 1
                            End If
                        Case "o":
                            If (Sheets(SheetName).Cells(k, OutputNameCol) = Name) And (OutDevice = "LOK") Then
                                ActiveSheet.Cells(j, LocalOutputAddressCol) = Sheets(SheetName).Cells(k, OutputAddressCol)
                            ElseIf (Sheets(SheetName).Cells(k, OutputNameCol) = Name) And (OutDevice = "STP") Then
                                ActiveSheet.Cells(j, LocalStopAddressCol) = Sheets(SheetName).Cells(k, OutputAddressCol)
                            ElseIf (Sheets(SheetName).Cells(k, OutputNameCol) = Name) And (OutDevice = "CLS") Then
                                ActiveSheet.Cells(j, LocalCloseAddressCol) = Sheets(SheetName).Cells(k, OutputAddressCol)
                            End If
                    End Select
                Next
            End If
        End If
        j = j + 1
    End If
Next

'Import standalone Outputs (non-LOK, non-STP, non-CLS Outputs) like ERM or ELEV that don't have any associated inputs
For i = IOStartRow To OutputNumOfObjects + IOStartRow - 1
    OutputDevice = Sheets(SheetName).Cells(i, OutputDeviceCol)
    If OutputDevice <> "SPR" And OutputDevice <> "LOK" And OutputDevice <> "STP" And OutputDevice <> "CLS" Then
        ActiveSheet.Cells(j, LocalInputDeviceCol) = Sheets(SheetName).Cells(i, OutputDeviceCol)  'Fill in the "DPS" cell
        ActiveSheet.Cells(j, LocalInputOutput) = "Lone Output"  'Fill in the Input/Ouput cell
        ActiveSheet.Cells(j, LocalNameCol) = Sheets(SheetName).Cells(i, OutputNameCol)   'Fill in the device name cell
        ActiveSheet.Cells(j, LocalSideCol) = Sheets(SheetName).Cells(i, OutputSideCol)   'fill in the H/S side cell
        ActiveSheet.Cells(j, LocalIndexCol) = "N/A"    'Fill in the "p 1" Indusoft Index cell with N/A since there should be no associated index
        ActiveSheet.Cells(j, LocalInputAddressCol) = Sheets(SheetName).Cells(i, OutputAddressCol)    'Fill in the input address cell
        ActiveSheet.Cells(j, LocalDeviceInfo) = Sheets(SheetName).Cells(i, OutputDeviceInfo)    'Fill in the output device info cell
        j = j + 1
    End If
Next

'Import outstanding Door Outputs - ie a LOK or STP or CLS without a corresponding DPS on the input side left there by accident
k = LocalStartRow
For i = IOStartRow To OutputNumOfObjects + IOStartRow - 1
    OutputDevice = Sheets(SheetName).Cells(i, OutputDeviceCol)  'LOK, STP, or CLS
    If OutputDevice = "LOK" Or OutputDevice = "STP" Or OutputDevice = "CLS" Then
        For k = IOStartRow To InputNumOfObjects 'Scan the input side names
            InDevice = Sheets(SheetName).Cells(k, InputDeviceCol)
            InName = Sheets(SheetName).Cells(k, InputNameCol)
            OutName = Sheets(SheetName).Cells(i, OutputNameCol)
            If InName = OutName And InDevice = "DPS" Then
                Match = 1
                k = InputNumOfObjects + 1
            Else: Match = 0
            End If
        Next
        If Match = 0 Then   'if there is no corresponding DPS input then
            ActiveSheet.Cells(j, LocalInputDeviceCol) = Sheets(SheetName).Cells(i, OutputDeviceCol)  'Fill in the "DPS" cell
            ActiveSheet.Cells(j, LocalInputOutput) = "Unmatched Output"  'Fill in the Input/Ouput cell
            ActiveSheet.Cells(j, LocalNameCol) = Sheets(SheetName).Cells(i, OutputNameCol)   'Fill in the device name cell
            ActiveSheet.Cells(j, LocalSideCol) = Sheets(SheetName).Cells(i, OutputSideCol)   'fill in the H/S side cell
            ActiveSheet.Cells(j, LocalIndexCol) = "N/A"    'Fill in the "p 1" Indusoft Index cell with N/A since there should be no associated index
            ActiveSheet.Cells(j, LocalInputAddressCol) = Sheets(SheetName).Cells(i, OutputAddressCol)    'Fill in the input address cell
            j = j + 1
        End If
    End If
Next
Set TargetRange = Range(Cells(1, LocalInputDeviceCol), Cells(j, LocalDeviceInfo))
TargetRange.Columns.AutoFit  'Fit the cells to the text
End Sub

Public Function OtherSheetEnumerate(SheetName, StartRow, Column)
Dim i, Tally
Tally = 0
For i = StartRow To 3000
    If Sheets(SheetName).Cells(i, Column) <> "" Then Tally = Tally + 1
Next
OtherSheetEnumerate = Tally
End Function

Public Sub ClearAll()
Dim i, j, StartCol, StartRow, NumOfEntries, EndColumn
StartRow = 1
EndColumn = 8
ActiveSheet.Cells.Range(Cells(1, 2), Cells(3000, 3000)).ClearContents
End Sub

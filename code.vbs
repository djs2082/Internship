============================================================================================
'Created by DILIP JOSHI
'Just Run this file in and it will generate V1 UTs report for all Teams and their TestCases.
'=============================================================================================
Set args = Wscript.Arguments
Dim excelFile
Dim sheetName
Dim src
Dim dest
Dim scriptdir
Dim vf_last_col
dtmStart = Now
'Create Excel and workbooks object
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set excelObj=CreateObject("Excel.Application")
excelObj.DisplayAlerts = False
excelFile="V1UnitTestResults.xlsx"
Set wrkBookObj=excelObj.WorkBooks.open(scriptdir+"\"+excelFile)

Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile scriptdir+"\"+excelFile,scriptdir+"\"+"V1UnitTestResults-backup.xlsx",True
'Wscript.Echo "Entering the Entries into DONOTDELETE sheet...."
'Enter the Entries into DoNotDelete sheet
sheetName="Version_History"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
keyvalue=sheetObj.Range("G6").value
''MsgBox keyvalue
lc=sheetObj.Range(keyvalue&"1").Column
''MsgBox lc
kzug=Split(sheetObj.Cells(, lc-3).Address, "$")(1)
lc=lc+4
destRangeStart=Split(sheetObj.Cells(, lc).Address, "$")(1)
''MsgBox destRangeStart
sheetObj.Range("G6").value=destRangeStart
sheetName="DoNotDelete"



Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile ("D:\V2Framework\V2Setup\SetEnvironment.bat", 1)
Do Until file.AtEndOfStream
  line = file.Readline
  if(Left(line,11)="SET BUILDID") Then
  newBuild=Right(line,19)
  End if
Loop
file.Close


' newBuild=InputBox("Enter New Global Build Version Name")








Set sheetObj=wrkBookObj.Worksheets(sheetName)
Last_Row_DND=DoNotDelete(wrkBookObj,sheetObj,newBuild)
'Wscript.Echo "Creating the Entries into CONSOLIDATED RESULT sheet...."
'Enter the Entries into Consolidated_Result sheet
sheetName="Consolidated Result"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Consolidated_Result wrkBookObj,sheetObj,Last_Row_DND
Create_bar_graph wrkBookObj,sheetObj,Last_Row_DND
'Wscript.Echo "Editing the bar-graph into CONSOLIDATED RESULT sheet...."
Create_line_graph wrkBookObj,sheetObj,Last_Row_DND
'Wscript.Echo "Editing the line-graph into CONSOLIDATED RESULT sheet...."
Last_Column=sheetObj.UsedRange.Columns.Count

'Enter the entries into BGR3 sheet

sheetName="BGR3"
Team_no=1
Last_Col=5
Source_Range="PO1:PQ6"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
' ck=sheetObj.UsedRange.Columns.Count
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
 ck=sheetObj.UsedRange.Columns.Count
'Wscript.Echo "Creating the Entries into BGR3 sheet...."
'Enter the entries into INT1 sheet
sheetName="INT1"
Team_no=2
Last_Col=11
Source_Range="PO1:PQ11"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into INT1 sheet...."
'Enter the entries into MIL4 sheet
sheetName="MIL4"
Team_no=3
Last_Col=29
Source_Range="PO1:PQ29"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into MIL4 sheet...."
'Enter the entries into MIL2 sheet
sheetName="MIL2"
Team_no=4
Last_Col=31
Source_Range="PO1:PQ31"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into MIL2 sheet...."
'Enter the entries into ZUG1 sheet
sheetName="ZUG1"
Team_no=5
Last_Col=11
Source_Range="PO1:PQ11"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into ZUG1 sheet...."
'Enter the entries into MIL1 sheet
sheetName="MIL1"
Team_no=6
Last_Col=30
Source_Range="PO1:PQ30"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into MIL1 sheet...."
'Enter the entries into PUN1 sheet
sheetName="PUN1"
Team_no=7
Last_Col=6
Source_Range="PO1:PQ6"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into PUN1 sheet...."
'Enter the entries into BGR1 sheet
sheetName="BGR1"
Team_no=8
Last_Col=17
Source_Range="PO1:PQ17"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into BGR1 sheet...."
'Enter the entries into PUN3 sheet
sheetName="PUN3"
Team_no=9
Last_Col=23
Source_Range="PO1:PQ23"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into PUN3 sheet...."
'Enter the entries into PUN3 sheet
sheetName="BANA2"
Team_no=10
Last_Col=17
Source_Range="PO1:PQ17"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into bana2 Sheet...."
'Enter the Entries into MIL1-Scripting sheet
sheetName="MIL1-Scripting"
Team_no=11
Last_Col=6
Source_Range="PO1:PQ6"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into MIL1-Scripting sheet...."
'Enter the Entries into Video sheet
sheetName="Video"
Source_Range="PO1:PQ16"
Team_no=12
Last_Col=16
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into VIDEO sheet...."
'Enter the Entries into Siclimat sheet
sheetName="Siclimat"
Source_Range="PO1:PQ6"
Team_no=13
Last_Col=6
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into Siclimat sheet...."
'Enter the Entries into Pun2 sheet
sheetName="PUN2"
Team_no=14
Last_Col=10
Source_Range="PO1:PQ10"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo"Creating the Entries into PUN2 sheet...."
'Enter the Entries into WSI(ZUG2) sheet
sheetName="WSI(ZUG2)"
Team_no=15
Last_Col=4
Source_Range="PO1:PQ4"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo "Creating the Entries into WSI(ZUG2) sheet...."

'Enter the Entries into BANA1 sheet
sheetName="BANA1"
Team_no=16
Last_Col=4
Source_Range="PO1:PQ4"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo "Creating the Entries into BANA1 sheet...."
'Enter the Entries into ZUG5-BAEU sheet
sheetName="ZUG5-BAEU"
Team_no=17
Last_Col=6
Source_Range="PL1:PN6"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,kzug
'Wscript.Echo "Creating the Entries into zug5-Baeu sheet...."
'Enter the Entries into UnSpecified sheet
sheetName="UnSpecified"
Team_no=18
Last_Col=19
Source_Range="PO1:PQ19"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
Teams wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue
'Wscript.Echo "Creating the Entries into Unspecified sheet...."
'Enter the new blank graph in Teamwise graph sheet
sheetName="Teamwise Graphs"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
NewGraph wrkBookObj,sheetObj,Last_Column,newBuild,Last_Row_DND
'Wscript.Echo "Creating the new teamwise graph...."
sheetName="BGR3"
' Set sheetObj=wrkBookObj.Worksheets(sheetName)
' ck=sheetObj.UsedRange.Columns.Count
excelObj.DisplayAlerts = False
wrkBookObj.Save
wrkBookObj.Close True
excelObj.Quit
'Wscript.Echo "sleep 2 sec...."
' Wscript.Sleep(2000)
' scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
' Set excelObj=CreateObject("Excel.Application")
' excelObj.DisplayAlerts = False
' excelFile="V1UnitTestResults.xlsx"
' Set wrkBookObj=excelObj.WorkBooks.open(scriptdir+"\"+excelFile)
' sheetName="BGR3"
' Set sheetObj=wrkBookObj.Worksheets(sheetName)
' ck=sheetObj.UsedRange.Columns.Count
 Set oShell = CreateObject("WScript.Shell")
' Wscript.Echo "Running Update_V1_UT_Count.bat file...."
 oShell.CurrentDirectory = "D:\V2Framework\V2Setup\Update"
 command="Update_V1_UT_Count.bat"
 ShowWindow=0
 WaitUntilFinished=true
 oShell.Run command, ShowWindow, WaitUntilFinished
 
 
 ' ck=sheetObj.UsedRange.Columns.Count
 
 
' oShell.run "cmd.exe D:\V2Framework\V2Setup\Update\Update_V1_UT_Count.bat"
' Thread.Sleep 20000
'Wscript.Echo "sleep 2 sec...."
Wscript.Sleep(2000)


scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set excelObj=CreateObject("Excel.Application")
excelObj.DisplayAlerts = False
excelFile="V1UnitTestResults.xlsx"
Set wrkBookObj=excelObj.WorkBooks.open(scriptdir+"\"+excelFile)



sheetName="BGR3"
Set sheetObj=wrkBookObj.Worksheets(sheetName)
vf_last_col=sheetObj.UsedRange.Columns.Count
' MsgBox vf_last_col
' MsgBox ck
if vf_last_col=ck+3 Then
' Wscript.Echo "results are not placed at their proper position...."
' Wscript.Echo "So Arranging the results at their position...."
sheetName="BGR3"
Last_Col=5
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="INT1"
Last_Col=11
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="MIL4"
Last_Col=29
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="MIL2"
Last_Col=31
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="ZUG1"
Last_Col=11
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="MIL1"
Last_Col=30
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="PUN1"
Last_Col=6
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="BGR1"
Last_Col=17
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="PUN3"
Last_Col=23
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="BANA2"
Last_Col=17
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="MIL1-Scripting"
Last_Col=6
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="Video"
Last_Col=16
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="Siclimat"
Last_Col=6
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="PUN2"
Last_Col=10
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="WSI(ZUG2)"
Last_Col=4
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="BANA1"
Last_Col=4
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="ZUG5-BAEU"
Last_Col=6
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

sheetName="UnSpecified"
Last_Col=19
Set sheetObj=wrkBookObj.Worksheets(sheetName)
correctsheet sheetObj,Last_Col

end if

' Save And Close the Sheet
excelObj.DisplayAlerts = False
wrkBookObj.Save
wrkBookObj.Close True
excelObj.Quit

'Greetings
dtmEnd = Now
MsgBox "Report is created in " & DateDiff("s", dtmStart, dtmEnd) & " seconds"

'Quit the Script
Wscript.Quit



'==============================================================================================
'FUNCTION DoNotDelete
'==============================================================================================



Function DoNotDelete(wrkBookObj,sheetObj,newBuild)
Last_Row=sheetObj.UsedRange.Rows.Count
Last_Row=Last_Row+1

'create destination range
destRange= "B" & Last_Row+1& ":" & "D" & Last_Row+1
Last_Row_Value=sheetObj.Cells(Last_Row,2).Value

'copy row from source(B4:D4) at destinataion
sheetObj.Range("B4:D4").Copy sheetObj.Range(destRange)

'enter values in copied row
sheetObj.Cells(Last_Row+1,2).Value=Last_Row_Value+1
sheetObj.Cells(Last_Row+1,3).Value = newBuild
todays_date=CStr(FormatDateTime(Now(),vbLongDate))
sheetObj.Cells(Last_Row+1,4).Value=todays_date
sheetObj.Cells(Last_Row+1,4).HorizontalAlignment = -4131
DoNotDelete=Last_Row

end Function



'=======================================================================================================
'SUBROUTINE Consolidated_Result
'=======================================================================================================



sub Consolidated_Result(wrkBookObj,sheetObj,Last_Row_DND)
Dim Last_Row
Dim Last_Column
Dim destRange

Last_Column=sheetObj.UsedRange.Columns.Count
Last_Row=sheetObj.UsedRange.Rows.Count
'Get the Destination Range
destRangeStart=Split(sheetObj.Cells(, Last_Column+1).Address, "$")(1)
rCcnt=sheetObj.Range("B1:E22").Columns.Count
rRcnt=sheetObj.Range("B1:E22").Rows.Count
destRangeEnd=Split(sheetObj.Cells(, Last_Column+rCcnt).Address, "$")(1)
destRange=destRangeStart & "1:" & destRangeEnd & rRcnt

'copy and paste table from Source(F1:I22) to Destination
sheetObj.Range("ET1:EW22").Copy sheetObj.Range(destRange)

'initialize the copied table
For iR = 4 to 21
For iC = Last_Column+1 to Last_Column+3
sheetObj.Cells(iR,iC)=0
Next
Next

'Add the Formulas to copied Table
formula="=DoNotDelete!" & "D" & Last_Row_DND+1
sheetObj.Range(destRangeStart&"1").Formula=formula
formula="=DoNotDelete!" & "C" & Last_Row_DND+1
sheetObj.Range(destRangeStart&"2").Formula=formula

 end sub
 
 
 
'====================================================================================================
'SUBROUTINE TEAMS
'=====================================================================================================



sub Teams(wrkBookObj,sheetObj,Last_Row_DND,Source_Range,Team_no,Last_Col,keyvalue)
Dim Last_Row
Dim Last_Column
Dim destRange
sheetObj.Rows(Last_Col+1 & ":" & sheetObj.Rows.Count).Delete
fn=keyvalue & "1" & ":" & "XFD" & Last_Col
''MsgBox fn
sheetObj.Range(fn).clear
Last_Column=sheetObj.UsedRange.Columns.Count
Last_Row=sheetObj.UsedRange.Rows.Count
' ''MsgBox Last_Column
' ''MsgBox Last_Row
'Get the Destination Range
destRangeStart=Split(sheetObj.Cells(, Last_Column+1).Address, "$")(1)
rCcnt=sheetObj.Range(Source_Range).Columns.Count
rRcnt=sheetObj.Range(Source_Range).Rows.Count
destRangeEnd=Split(sheetObj.Cells(, Last_Column+rCcnt).Address, "$")(1)
destRange=destRangeStart & "1:" & destRangeEnd & rRcnt

'Copy the Table from 'Source_Range' to Destination
sheetObj.Range(Source_Range).Copy sheetObj.Range(destRange)

'Intitialize the copied Table
For iR = 3 to Last_Col-1
For iC = Last_Column+1 to Last_Column+3
sheetObj.Cells(iR,iC)=0
Next
Next

'Add the Formulas to copied Table
formula="=DoNotDelete!" & "C" & Last_Row_DND+1
sheetObj.Range(destRangeStart&"1").Formula=formula

'Provide Mapping to Consolidated_Result sheet
sheetName="Consolidated Result"
Set sheetObj2=wrkBookObj.Worksheets(sheetName)
name="A" & Team_no+3
Lc=sheetObj2.UsedRange.Columns.Count

destRange=Split(sheetObj.Cells(, Last_Column+3).Address, "$")(1)
Last_Column_CR_let=destRange & "$" & Last_Col
sheetObj2.Cells(Team_no+3,Lc-1).Formula= "='" & sheetObj2.Range(name).value &"'!" & Last_Column_CR_let

destRange=Split(sheetObj.Cells(, Last_Column+2).Address, "$")(1)
Last_Column_CR_let=destRange & "$" & Last_Col
sheetObj2.Cells(Team_no+3,Lc-2).Formula= "='" & sheetObj2.Range(name).value &"'!" & Last_Column_CR_let

destRange=Split(sheetObj.Cells(, Last_Column+1).Address, "$")(1)
Last_Column_CR_let=destRange & "$" & Last_Col
sheetObj2.Cells(Team_no+3,Lc-3).Formula="='" & sheetObj2.Range(name).value &"'!" & Last_Column_CR_let

end sub



'============================================================================
'SUBROUTINE Create_bar_graph
'===========================================================================================



sub Create_bar_graph(wrkBookObj,sheetObj,Last_Row_DND)
chartName="Chart 1"
Last_Column=sheetObj.UsedRange.Columns.Count

'New Series for 'passed'
destRange1=Split(sheetObj.Cells(, Last_Column-3).Address, "$")(1)
destRange1="$" & destRange1 & "$22),1)"

'New Series for 'failed'
destRange2=Split(sheetObj.Cells(, Last_Column-2).Address, "$")(1)
destRange2="$" & destRange2 & "$22),2)"

'New Series for 'not run'
destRange3=Split(sheetObj.Cells(, Last_Column-1).Address, "$")(1)
destRange3="$" & destRange3 & "$22),3)"


set oChart=sheetObj.ChartObjects("Chart 1")

'Assign new Series for 'passed'
set mySrs1=oChart.Chart.SeriesCollection(1)
OldString1=mySrs1.Formula
OldString1=Left(OldString1,Len(OldString1)-4)
OldString1=Right(OldString1,Len(OldString1)-40)
NewString1="=SERIES(""Passed""" & ",DoNotDelete!$C$4:$C$"& Last_Row_DND+1 & OldString1 & ",'Consolidated Result'!" & destRange1
mySrs1.Formula=NewString1

'Assign new Series for 'failed'
set mySrs2=oChart.Chart.SeriesCollection(2)
OldString2=mySrs2.Formula
OldString2=Left(OldString2,Len(OldString2)-4)
NewString2=OldString2 & ",'Consolidated Result'!" & destRange2
mySrs2.Formula=NewString2

'Assign new Series for 'not run'
set mySrs3=oChart.Chart.SeriesCollection(3)
OldString3=mySrs3.Formula
OldString3=Left(OldString3,Len(OldString3)-4)
NewString3=OldString3 & ",'Consolidated Result'!" & destRange3
mySrs3.Formula=NewString3

end sub



'===========================================================================================
'SUBROUTINE Create_line_graph
'=============================================================================================



sub Create_line_graph(wrkBookObj,sheetObj,Last_Row_DND)
chartName="Chart 2"
Last_Column=sheetObj.UsedRange.Columns.Count

'New Series for 'passed'
destRange1=Split(sheetObj.Cells(, Last_Column-3).Address, "$")(1)
destRange1="$" & destRange1 & "$22),1)"

'New Series for 'failed'
destRange2=Split(sheetObj.Cells(, Last_Column-2).Address, "$")(1)
destRange2="$" & destRange2 & "$22),2)"

set oChart=sheetObj.ChartObjects("Chart 2")

'Assign new Series for 'passed'
set mySrs1=oChart.Chart.SeriesCollection(1)
OldString1=mySrs1.Formula
OldString1=Left(OldString1,Len(OldString1)-4)
OldString1=Right(OldString1,Len(OldString1)-40)
NewString1="=SERIES(""Passed""" & ",DoNotDelete!$C$4:$C$"& Last_Row_DND+1 & OldString1 & ",'Consolidated Result'!" & destRange1
mySrs1.Formula=NewString1

'Assign new Series for 'failed'
set mySrs2=oChart.Chart.SeriesCollection(2)
OldString2=mySrs2.Formula
OldString2=Left(OldString2,Len(OldString2)-4)
NewString2=OldString2 & ",'Consolidated Result'!" & destRange2
mySrs2.Formula=NewString2

end sub



'==================================================================================
'SUBROUTINE NewGraph
'=====================================================================================


sub NewGraph(wrkBookObj,sheetObj,Last_Column_Consolidated,newBuild,Last_Row_DND)
Last_Row=sheetObj.UsedRange.Rows.Count
Last_Row=Last_Row+27

'Get Destination Range for title
destRange1=Split(sheetObj.Cells(, Last_Column_Consolidated-3).Address, "$")(1)
destRange2=Split(sheetObj.Cells(, Last_Column_Consolidated-2).Address, "$")(1)
destRange1="$" & destRange1 & "$4:$" & destRange1 & "21,1)"
destRange2="$" & destRange2 & "$4:$" & destRange2 & "21,2)"
destination="B" & Last_Row

'Assign Title to Chart
sheetObj.Range(destination).Formula="=DoNotDelete!C" & Last_Row_DND+1

'Get Destination Range for Chart
Last_Row=Last_Row+1
destination="B" & Last_Row

'Copy the Chart from source(B4028:k4064) to distination
sheetObj.Range("B4028:K4054").Copy sheetObj.Range(destination)

'Get the name of Newly Copied Chart
For Each ChtObj In sheetObj.ChartObjects
chartName=ChtObj.Name
Next

'Assing the formula for 'passed' to copied chart
set oChart=sheetObj.ChartObjects(chartName)
set mySrs1=oChart.Chart.SeriesCollection(1)
OldString1=mySrs1.Formula
OldString1=Left(OldString1,Len(OldString1)-15)
NewString1=OldString1 & destRange1
mySrs1.Formula=NewString1

'Assing the formula for 'passed' to copied chart
set mySrs2=oChart.Chart.SeriesCollection(2)
OldString2=mySrs2.Formula
OldString2=Left(OldString2,Len(OldString2)-15)
NewString2=OldString2 & destRange2
mySrs2.Formula=NewString2

end sub


'=================================================================================
'Correct Sheet
'==================================================================================


sub correctsheet(sheetObj,Last_Col)
Dim Last_Row
Dim Last_Column
Dim destRange

Last_Column=sheetObj.UsedRange.Columns.Count
Last_Row=sheetObj.UsedRange.Rows.Count


rCcnt=sheetObj.Range(Source_Range).Columns.Count
rRcnt=sheetObj.Range(Source_Range).Rows.Count

destRangeStart=Split(sheetObj.Cells(, Last_Column-2).Address, "$")(1)
' ''MsgBox destRangeStart

' ''MsgBox Last_Col
' ''MsgBox Last_Column-2
For iR = 3 to Last_Col-1
For iC = Last_Column-2 to Last_Column
sheetObj.Cells(iR,iC-3)=sheetObj.Cells(iR,iC)
Next
Next

' ''MsgBox destRangeStart & "1:XFD" & Last_Col
sheetObj.Rows(Last_Col+1 & ":" & sheetObj.Rows.Count).Delete
sheetObj.Range(destRangeStart & "1:XFD" & Last_Col+1).clear
End Sub


'=================================================================================
'Correct Sheet
'==================================================================================
'=================================================================================
'THE END
'==================================================================================

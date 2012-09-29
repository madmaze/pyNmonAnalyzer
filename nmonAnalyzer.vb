' THIS SOURCE IS EXTRACTED FROM THE ORIGINAL EXCEL MACRO
' WRITTEN BY IBM (https://www.ibm.com/developerworks/aix/library/au-nmon_analyser/)
' This is reference material for deciphering NMONs output

Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
                                       'last mod v3.3.h
Option Explicit
Public Analyser As String              'Analyser version (from main sheet)
Public Batch As Integer                'Batch mode setting (0/1)
Public BBBFont As String               'name of fixed pitch font to use on BBB sheets
Public ColNum As Integer               'Last column number
Public Copies As Integer               'Number of copies to print
Public CPUrows As Long                 'number of rows on the first CPU sheet
Public CPUmax As Integer               'maximum number of CPU/PCPU/SCPU sheets to create
Public DebugVar As Variant             'useful for debugging
Public Delim As String                 'Delimiter used in .csv file
Public DecSep As String                'Decimalseparator used in .csv file
Public dLPAR As Variant                'True if dynamic LPAR has been used for CPUs
Public ErrMsg As String                'Error message for file
Public Esoteric As String              'Either "hdiskpower" or "dac"
Public ESSanal As Boolean              'True if ESS analysis to be done
Public Filename As String              'Qualified name of the input file
Public First As Long                   'First time interval to process
Public FirstTime As Date               'First time/date to process
Public GotEMC As Boolean               'True if either EMC or FAStT present
Public GotESS As Boolean               'True if EMC/ESS/FAStT or DGs present
Public Graphs As String                'ALL/LIST
Public Host As String                  'Hostname from AAA sheet
Public Last As Long                    'Last time interval to process
Public LastTime As Date                'Last time/date to process
Public LastColumn As String            'Last column letter
Public Linux As Boolean                'True if Linux
Public List As String                  'List of sheets to graph
Public LScape As Boolean               'Value of LScape field
Public MaxRows As Long                 'Maximum rows in a sheet
Public Merge As String                 'NO, YES or KEEP
Public NumCPUskipped As Integer        'Number of CPU sections skipped
Public progname As String              'Version of NMON/topas
Public NoList As Boolean               'true if NOLIST=DELETE
Public NoTop As Boolean                'True if TOP section is to be deleted
Public NumCPUs As Integer              'Number of CPU sections
Public NumDisk As Integer              'Number of disk subsections
Public OutDir As String                'Name of output directory
Public Output As String                'CHART/PICTURES/WEB/PRINT
Public Pivot As Boolean                'True if a Pivot Chart is to be produced
Public PNG As Variant                  'AllowPNG (True/False)
Public Printer As String               'Name of Printer
Public Reorder As Boolean              'Reorder sheets after analysis (True/False)
Public rH As Single                    'Row Height
Public ReProc As Boolean               'Reprocess or bypass input files in batch mode
Public RunDate As String               'NMON run date from AAA sheet
Public SMTon As Boolean                'true if SMT is on (set in BBBP)
Public SMTmode As Integer              'number of threads per core
Public Snapshots As Long               'Set by PP_AAA and reset by PP_ZZZZ
Public SortInp As Boolean              'Sort input file (True/False)
Public Start As Double                 'Start date/time value
Public SubDir As Boolean               'OrganizeInFolder (True/False)
Public SVCTimes As Boolean             'Produce disk service time estimates (True/False)
Public SVCXLIM As Integer              'Lower limit for service time analysis
Public ThouSep As String               'Thousands separator unsed in .csv file
Public topas As Variant                'True if topas (PTX:xmtrend)
Public TopDisks As Integer             'No. of entries to show on graphs (0 = all)
Public TopRows As Long                 'No. of rows on the TOP sheet for PP_UARG
Public t1 As String                    'First timestamp on the CPU_ALL sheet
Public WebDir As String                'Name of output directory for HTML
Public xToD As String                  'Number format for ToD graphs
                                       'position & dimension data for charts
Public cTop As Integer, cLeft As Integer, cWidth As Integer, cHeight As Integer
Public csTop1 As Integer, csTop2 As Integer, csTop3 As Integer, csTop4 As Integer
Public csHeight As Integer
                                       'definitions for Pivot chart
Public PivotParms As Variant
                                       'definitions for SORT process
#If VBA7 Then
Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
       ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
       ByVal dwProcessId As Long) As Long
Public Declare PtrSafe Function WaitForSingleObject Lib "kernel32" ( _
       ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
       ByVal hObject As Long) As Long
#Else
Public Declare Function OpenProcess Lib "kernel32" ( _
       ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
       ByVal dwProcessId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" ( _
       ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" ( _
       ByVal hObject As Long) As Long
#End If

Public Const SYNCHRONIZE = &H100000
Public Const INFINITE = &HFFFF
                                       'last mod v3.3.e3
Sub ApplyStyle(Chart1 As ChartObject, gtype As Integer, nseries As Integer)
	'gtype = 0 for bars, 1 for lines, 2 for area
	'nseries = number of series
	Dim CurCol As Integer
	Dim ic As Integer
	Dim MaxVal As Double
	Dim Pct As String
	'Public xToD as string                  'Number format for ToD graphs
	    
	With Chart1.Chart
	   .ChartArea.AutoScaleFont = False
	   Chart1.Placement = xlFreeFloating
	   If .HasLegend = True Then .Legend.Position = xlLegendPositionTop
	   If .Axes(xlCategory).TickLabels.Orientation <> xlHorizontal Then _
	      .Axes(xlCategory).TickLabels.Orientation = xlUpward
	   If gtype = 1 Then
	      .Axes(xlValue).HasMajorGridlines = True
	      If nseries < 8 Then
		 For ic = 1 To nseries
		    .SeriesCollection(ic).Border.Weight = xlMedium
		 Next
	      End If
	   Else
	      If gtype = 0 Then
		 .ChartType = xlColumnStacked
		 .Axes(xlCategory).TickLabelSpacing = 1
	      Else
	      '  .Type = $Area
	      End If
	   End If
	   If gtype > 0 Then
	      .Axes(xlCategory).HasMajorGridlines = False
	      .Axes(xlCategory).MajorTickMark = xlNone
	      .Axes(xlCategory, xlPrimary).CategoryType = xlCategoryScale
	      .Axes(xlCategory).TickLabels.NumberFormat = xToD
	   End If
	      .Axes(xlValue).MinimumScale = 0
	End With
						'scale the y-axis
	With Chart1.Chart
	   MaxVal = .Axes(xlValue).MaximumScale
	   If .Axes(xlValue).DisplayUnit = xlThousands Then MaxVal = MaxVal / 1000
	   If MaxVal > 5181 Then
	      .Axes(xlValue).HasDisplayUnitLabel = True
	      If MaxVal > 3333431 Then
		 .Axes(xlValue).DisplayUnit = xlMillions
	      Else
		 .Axes(xlValue).DisplayUnit = xlThousands
	      End If
	   End If
	   Pct = ""
	   If Right(.Axes(xlValue).TickLabels.NumberFormat, 1) = "%" Then
	      Pct = "%"
	      MaxVal = MaxVal * 10
	   End If
	   If MaxVal >= 10 Then
	      .Axes(xlValue).TickLabels.NumberFormat = "0" & Pct
	   Else
	      .Axes(xlValue).TickLabels.NumberFormat = "0.0" & Pct
	   End If
	End With

End Sub
                                       'last mod v3.0
Function avgmax(numrows As Long, Sheet1 As Worksheet, DoSort As Integer) As Range
	Dim Chart1 As ChartObject
	Dim Column As String
	Dim DiskData As Range
	'Public LastColumn As String, ColNum as Integer
	Dim MyCells As Range
	Dim n As Integer
						' Put in the formulas for avg/max
	Set avgmax = Sheet1.Range("A" & CStr(numrows + 2) & ":" & LastColumn & CStr(numrows + 4 + DoSort))
	Column = "B2:B" & CStr(numrows)
	avgmax.Item(1, 1) = "Avg."
	avgmax.Item(1, 2) = "=AVERAGE(" & Column & ")"
	avgmax.Item(2, 1) = "WAvg."
	avgmax.Item(2, 2) = "=IF(B" & CStr(numrows + 2) & "=0,0,MAX(SUMPRODUCT(" & Column & "," & Column & ")/SUM(" & Column & ")-B" & CStr(numrows + 2) & ",0))"
	avgmax.Item(3, 1) = "Max."
	avgmax.Item(3, 2) = "=ABS(MAX(B2:B" & CStr(numrows) & ")-B" & CStr(numrows + 2) & "-B" & CStr(numrows + 3) & ")"
	If DoSort = 1 Then
	   avgmax.Item(4, 1) = "SortKey"
	   avgmax.Item(4, 2) = "=B" & CStr(numrows + 2) & "+ B" & CStr(numrows + 3)
	End If
	
	If LastColumn <> "B" Then
	   Set MyCells = Sheet1.Range("B" & CStr(numrows + 2) & ":" & LastColumn & CStr(numrows + 4 + DoSort))
	   MyCells.FillRight
	   MyCells.NumberFormat = "0.0"
	End If
	      
End Function
Public Function Checklist(inVal As String) As Boolean
	Dim MyArray As Variant
	Dim i As Long
	
	Checklist = True
	MyArray = Split(List, ",")
	For i = 0 To UBound(MyArray)
	    If inVal Like MyArray(i) Then Exit Function
	Next i
	Checklist = False

End Function
Public Function ConvertRef(inVal As Variant) As Variant
	If IsNumeric(inVal) Then
	   If inVal < 26 Then
	      ConvertRef = Chr(inVal + 65)
	   Else
	      ConvertRef = Chr((inVal \ 26) + 64) & Chr((inVal Mod 26) + 65)
	   End If
	Else
	   If Len(inVal) > 1 Then
	      ConvertRef = ((Asc(UCase$(inVal)) - 64) * 26) + (Asc(UCase$(Right$(inVal, 1))) - 64)
	   Else
	      ConvertRef = Asc(UCase$(inVal)) - 64
	   End If
	End If
End Function
                                       'v3.3.g1
Sub CreatePivot()
	Dim i As Long, nr As Long, nc As Integer
	Dim MyCells As Range
	Dim pSheet As String
	Dim pPage As String
	Dim pRow As String
	Dim pColumn As String
	Dim pData As String
	Dim pFunc As Integer
	Dim Sheet1 As Worksheet
	
	pSheet = PivotParms(0)
	If Not SheetExists(pSheet) Then Exit Sub
	UserForm1.Label1.Caption = "Creating Pivot Chart for " & pSheet
	UserForm1.Repaint
	
	Set Sheet1 = Worksheets(pSheet)
	pPage = PivotParms(1)
	pRow = PivotParms(2)
	pColumn = PivotParms(3)
	pData = PivotParms(4)
	
	Select Case PivotParms(5)
	   Case ("SUM")
	      pFunc = -4157
	   Case ("COUNT")
	      pFunc = -4112
	   Case ("MIN")
	      pFunc = -4139
	   Case ("AVG")
	      pFunc = -4106
	   Case ("MAX")
	      pFunc = -4136
	   Case Else
	      pFunc = 1000
	End Select
					       'number of used columns
	For nc = 1 To 255
	   If Sheet1.Cells(1, nc) = "" Then Exit For
	Next
					       'number of used rows
	For nr = 1 To MaxRows
	   If Sheet1.Cells(nr, 1) = "" Then Exit For
	Next
					       'produce the pivot chart
	i = Worksheets("AAA").Range("snapshots")
	ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:= _
	   pSheet & "!A1:" & ConvertRef(nc - 2) & CStr(nr - 1)).CreatePivotTable TableDestination:="", _
	   TableName:="MyPivot"
	ActiveSheet.PivotTableWizard TableDestination:=ActiveSheet.Cells(3, 1)
	ActiveSheet.Cells(3, 1).Select
	ActiveSheet.PivotTables("MyPivot").AddFields RowFields:=pRow, _
	   ColumnFields:=pColumn, PageFields:=pPage
	With ActiveSheet.PivotTables("MyPivot").PivotFields(pData)
	   .Orientation = xlDataField
	   .Function = pFunc
	End With
	Charts.Add
	ActiveChart.Location Where:=xlLocationAsNewSheet
	ActiveChart.PlotArea.Select
	ActiveChart.ChartType = xlAreaStacked
	With ActiveChart.Axes(xlCategory)
	   .CrossesAt = 1
	   If pRow = "Time" And i > 10 Then .TickLabelSpacing = (i / 10)
	   .TickMarkSpacing = 1
	   .AxisBetweenCategories = False
	   .ReversePlotOrder = False
	End With
	ActiveSheet.Move After:=Sheets("AAA")
End Sub
                                       'v3.3.0
Sub DefineStyles()
	Dim aa0 As String
	'Public cTop As Integer, cLeft As Integer, cWidth As Integer, cHeight As Integer
	'Public csTop1 As Integer, csTop2 As Integer, csHeight As Integer
	Dim temp As Single
	
	If cWidth = 0 Then
	   If Output = "PRINT" Then
	      If LScape Then Worksheets(1).PageSetup.Orientation = xlLandscape
	      Worksheets(1).Cells(1, 255) = "break"
	      ActiveWindow.View = xlPageBreakPreview
	      aa0 = Worksheets(1).VPageBreaks(1).Location.Address(True, True, xlR1C1)
	      cWidth = (Right(aa0, Len(aa0) - InStr(1, aa0, "C")) - 1) * Worksheets(1).Range("A1").Width
	      ActiveWindow.View = xlNormalView
	      Worksheets(1).Cells(1, 255) = ""
	   Else
	      cWidth = Application.UsableWidth - (1.7 * Worksheets(1).Range("A1").Width)
	   End If
	   cHeight = Int(cWidth / 2.4)
	   temp = Worksheets(1).Range("A1").Height
	   cHeight = Int(Int(cHeight / temp) * temp)
	End If
						'position & dimensions for single charts
	rH = Worksheets(1).Range("A1").Height
	cTop = Worksheets(1).Range("A1").Height + 1
	cLeft = Worksheets(1).Range("A1").Width
						'position & dimensions for multi-charts
	csHeight = cHeight
	csTop1 = cTop
	csTop2 = csHeight + csTop1 + 1
	csTop3 = csHeight + csTop2 + 1
	csTop4 = csHeight + csTop3 + 1
End Sub
Sub DelInt(Sheet1 As Worksheet, numrows As Long, HdLines As Integer)
	If Last < numrows - HdLines Then
	   Sheet1.Range("A" & CStr(Last + HdLines + 1) & ":A" & CStr(numrows)).EntireRow.Delete
	   numrows = Last + HdLines
	End If
	If First > 1 Then
	   Sheet1.Range("A" & CStr(HdLines + 1) & ":A" & CStr(First - 1 + HdLines)).EntireRow.Delete
	   numrows = numrows - First + 1
	End If
End Sub
                                       'last mod v3.3.F
Sub DiskGraphs(numrows As Long, SectionName As String, DoSort As Variant)
	Static Chart1 As ChartObject            'new chart object
	Static ChartTitle As String             'title for new chart
	'Public ColNum As Integer               'last column number
	Static Column As String                 'column of last disk to be graphed
	Static DiskData As Range                'data to be charted
	'Public Host As String                  'Hostname from AAA sheet
	'Public  LastColumn As String           'last column letter
	Static MyCells As Range                 'temp var
	Static MyWidth As Integer               'variable graph width
	Static nd As Integer                    'number of disks to graph
	'Public RunDate As String               'NMON run date from AAA sheet
	Static Sheet1 As Worksheet              'current sheet
	'Public TopDisks As Integer             'No. of hdisks to show on graphs (0 = all)
	    
	Set Sheet1 = Worksheets(SectionName)
	Sheet1.Activate
	'nd = InStr(1, List, Sheet1.Name)
						'Change column widths
	Set MyCells = Sheet1.Range("B1:" & LastColumn & CStr(numrows))
	MyCells.ColumnWidth = 7
	If Graphs = "LIST" And Not Checklist(Sheet1.Name) Then Exit Sub
						'preset chart title
	ChartTitle = Sheet1.Range("A1")
	If Left(SectionName, 4) = "DISK" And InStr(1, SectionName, "SIZE") > 0 Then
	   ChartTitle = ChartTitle + " (Kbytes)"
	End If
						 'Put in the formulas for avg/max
	Set DiskData = avgmax(numrows, Sheet1, 1)
	If DoSort Then
	   Sheet1.Range("B1:" & LastColumn & CStr(numrows + 5)).Sort _
	    Key1:=Sheet1.Range("B" & CStr(numrows + 5)), Order1:=xlDescending, _
	    Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlLeftToRight
	End If
	
	nd = DiskData.Columns.Count - 1
	If (nd > TopDisks And TopDisks > 0) Then
	   Column = ConvertRef(TopDisks)
	   ChartTitle = ChartTitle & " (1st " & CStr(TopDisks) & ")"
	   nd = TopDisks
	Else
	   Column = LastColumn
	End If
	
	MyWidth = cWidth / 50 * nd
	If MyWidth < cWidth Then MyWidth = cWidth
	Set DiskData = Sheet1.Range("B" & CStr(numrows + 2) & ":" & Column & CStr(numrows + 4))
	ChartTitle = ChartTitle & "  " & RunDate
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + rH * numrows, MyWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=DiskData, PlotBy:=xlRows, Title:=ChartTitle
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).Name = Sheet1.Cells(numrows + 2, 1).Value
	   .SeriesCollection(2).Name = Sheet1.Cells(numrows + 3, 1).Value
	   .SeriesCollection(3).Name = Sheet1.Cells(numrows + 4, 1).Value
	   .SeriesCollection(1).XValues = Sheet1.Range(Cells(1, 2), Cells(1, nd + 1))
						'apply customisation
	   If InStr(SectionName, "DISKBUSY") > 0 Then .Axes(xlValue).MaximumScale = 100
	   Call ApplyStyle(Chart1, 0, 3)
	End With
						'produce line graph
	Set DiskData = Sheet1.Range("A1:" & Column & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + rH * numrows, MyWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=DiskData, Gallery:=xlLine, Format:=2, _
	    Title:=ChartTitle, CategoryLabels:=1, SeriesLabels:=1, HasLegend:=True
						'apply customisation
	With Chart1.Chart
	    Call ApplyStyle(Chart1, 1, nd)
	End With

End Sub
                                        'v3.3.h
Sub GetIntervals(Filename As String)
	Dim fnin As Variant, fs As Variant
	Dim InputString As String, aa0 As Variant, temp As Long
	Dim TS As Double
	
	First = 0
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set fnin = fs.OpenTextFile(Filename, 1, 0)
						  'look for ZZZZ values
	Do While fnin.AtEndOfStream <> True
	   InputString = fnin.readline
	   If Left(InputString, 4) = "ZZZZ" Then
	      aa0 = Split(InputString, Delim)
	      temp = CLng(Right(aa0(1), Len(aa0(1)) - 1))
	      TS = TimeValue(aa0(2))
	      If First = 0 Then        'looking for first interval
		 If FirstTime > 1 Then TS = DateValue(aa0(3)) + TS  'looking for a date as well as a time
		 If TS >= FirstTime Then
		    First = temp
		    Last = First + 1
					'if LastTime is less than FirstTime then assume next day
		    If (LastTime < 1) And (FirstTime > LastTime) Then LastTime = LastTime + DateValue(aa0(3)) + 1
		 End If
	      Else                                'looking for last interval
		 If LastTime > 1 Then TS = DateValue(aa0(3)) + TS  'looking for a date as well as a time
		 If TS >= LastTime Then Exit Do
		 Last = temp
	      End If
	   End If
	Loop
	fnin.Close

End Sub
                                        'new for v2.9.0
Sub GetLastColumn(Sheet1 As Worksheet)
	'Public ColNum As Integer               'last column number
	'Public LastColumn As String            'last column letter
	Dim n As Integer
						'Locate the last column
	   LastColumn = "B"
	   For n = 3 To 255
	     If Sheet1.Range("A1")(1, n) = "" Then Exit For
	   Next n
	   ColNum = n - 1
	   LastColumn = Sheet1.Range("A1").Item(1, ColNum).Address(True, False, xlA1)
	   LastColumn = Left(LastColumn, InStr(1, LastColumn, "$") - 1)
End Sub
                                       'last mod v3.3.e
Sub GetNextSection(RawData As Worksheet, SectionName As String, StartRow As Long, CurrentRow As Long)
	'Public FileName As String             'Qualified name of the input file
	Dim aa1 As String * 1
	Dim InputString As String
	Dim j As Long                          'loop counter
	Dim skip As Variant                    'true if previous section was >65K
	
	If RawData.Cells(CurrentRow, 1) = SectionName Then
	   skip = True
	Else
	   SectionName = RawData.Cells(CurrentRow, 1).Value
	End If
	If SectionName = "" Then Exit Sub
	
	StartRow = CurrentRow
	If SectionName = RawData.Cells(MaxRows, 1) Then
	   If skip Then                         'skip remainder of >MaxRows section
	      UserForm1.Label1.Caption = "skipping remainder of " & SectionName
	      UserForm1.Repaint
	      RawData.UsedRange.Clear
	      Do
		 Line Input #1, InputString
		 If EOF(1) Then Close #1: SectionName = "": Exit Sub
	      Loop While Left(InputString, Len(SectionName)) = SectionName
	      RawData.Cells(1, 1).Value = InputString
	      CurrentRow = 1
	      StartRow = 1
	   End If
						'Part of section is on external file
						'delete processed lines
	   If CurrentRow > 1 Then
	      RawData.Range("A1:IV" & CStr(CurrentRow - 1)).Delete
	      CurrentRow = (MaxRows - StartRow + 1)
	      StartRow = CurrentRow + 1
	   End If
	      
	   If FreeFile() = 1 Then               'first time through we open file
	      UserForm1.Label1.Caption = "Repositioning file"
	      UserForm1.Repaint
	      Open Filename For Input As #1
						'and bypass first MaxRows lines
	      For j = 1 To MaxRows
		 Line Input #1, aa1
	      Next j
	   End If
	      
	   UserForm1.Label1.Caption = "Reading next section"
	   UserForm1.Repaint
	   Application.Calculation = xlCalculationManual
	   For CurrentRow = CurrentRow + 1 To MaxRows
	      Line Input #1, InputString
	      RawData.Cells(CurrentRow, 1).Value = InputString
	      If EOF(1) Then Exit For
	   Next CurrentRow
	   Application.Calculation = xlCalculationAutomatic
	   Application.Calculate
	   If EOF(1) Then
	      Close #1
	   Else
	      CurrentRow = CurrentRow - 1
	   End If
						   'parse the new lines
	   UserForm1.Label1.Caption = "Parsing section"
	   UserForm1.Repaint
	   
	   RawData.Range("A" & CStr(StartRow) & ":A" & CStr(CurrentRow)).TextToColumns , _
	      Destination:=Range("A" & CStr(StartRow)), DataType:=xlDelimited, _
	      Other:=True, OtherChar:=Delim, DecimalSeparator:=DecSep
	   
	   StartRow = 1
	   CurrentRow = 1
	   SectionName = RawData.Cells(CurrentRow, 1).Value
	End If
						'complete section is on the sheet
	Do
	   CurrentRow = CurrentRow + 1
	   If CurrentRow - StartRow = MaxRows - 1 Then Exit Do
	Loop While RawData.Cells(CurrentRow, 1).Value = SectionName

End Sub
                                       'last mod v3.3.h
Sub GetSettings(FileList As Variant, Numfiles As Integer)
	Dim Filename As String
	Dim Fname As String
	Dim fPath As String
	Dim i As Integer
	Dim MyCells As Range
	Dim sTemp As String
	Dim Sheet1 As Worksheet
	 
	Set Sheet1 = ThisWorkbook.Worksheets(1)
	Analyser = Sheet1.Range("A1").Value
					       'Settings for this run
	Graphs = Sheet1.Range("Graphs").Value
	First = Sheet1.Range("First").Value
	Last = Sheet1.Range("Last").Value
	FirstTime = Sheet1.Range("FirstTime").Value
	LastTime = Sheet1.Range("LastTime").Value
	Merge = Sheet1.Range("Merge").Value
	NoTop = Sheet1.Range("NoTop").Value = "NOTOP"
	Output = Sheet1.Range("Output").Value
	Pivot = Sheet1.Range("Pivot").Value = "YES"
	ESSanal = Sheet1.Range("ESSanal").Value = "YES"
	Filename = Sheet1.Range("Filelist").Value
	'================= Settings Sheet ======================================
	Set Sheet1 = ThisWorkbook.Worksheets("Settings")
					       'Batch settings
	ReProc = Sheet1.Range("Reproc").Value = "YES"
	OutDir = Sheet1.Range("OutDir").Value
	If OutDir <> "" Then
	   If Right(OutDir, 1) <> "\" Then OutDir = OutDir & "\"
	   If Dir(OutDir, vbDirectory) = "" Then
	      MsgBox ("Output Directory does not exist")
	      Exit Sub
	   End If
	End If
					       'Formatting settings
	BBBFont = Sheet1.Range("BBBFont").Value
	cWidth = Sheet1.Range("GWidth").Value
	cHeight = Sheet1.Range("GHeight").Value
	List = Sheet1.Range("List").Value + ",SYS_SUMM,CPU_SUMM,DISK_SUMM"
	CPUmax = Sheet1.Range("CPUmax").Value
	NoList = Sheet1.Range("NoList").Value = "DELETE"
	Reorder = Sheet1.Range("Reorder").Value = "YES"
	TopDisks = Sheet1.Range("TopDisks").Value
	xToD = Sheet1.Range("xToD").Value
					       'Pivot chart parameters
	If Pivot Then
	   PivotParms = Array(Sheet1.Range("PivotParms").Cells(1, 1), _
			      Sheet1.Range("PivotParms").Cells(1, 2), _
			      Sheet1.Range("PivotParms").Cells(1, 3), _
			      Sheet1.Range("PivotParms").Cells(1, 4), _
			      Sheet1.Range("PivotParms").Cells(1, 5), _
			      Sheet1.Range("PivotParms").Cells(1, 6))
	End If
					       'Printer settings
	LScape = Sheet1.Range("LScape").Value Like "YES"
	Copies = Sheet1.Range("Copies").Value
	Printer = Sheet1.Range("Printer").Value
					       'Web settings
	PNG = Sheet1.Range("PNG").Value Like "YES"
	SubDir = Sheet1.Range("SUBDIR").Value Like "YES"
	WebDir = Sheet1.Range("WebDir").Value
	If WebDir <> "" Then
	   If Right(WebDir, 1) <> "\" Then WebDir = WebDir & "\"
	   If Dir(WebDir, vbDirectory) = "" Then
	      MsgBox ("Web output Directory does not exist")
	      Exit Sub
	   End If
	End If
					       'National language settings
	Delim = Sheet1.Range("Delim").Value
	DecSep = ".": ThouSep = ","
	If Delim = ";" Then DecSep = ",": ThouSep = "."
	SortInp = Sheet1.Range("SortInp").Value Like "YES"
	'================= Build filelist ====================================
	Numfiles = 0
	Set MyCells = Worksheets(1).Range("FileList").Offset(0, -1)
	If Filename = "" Or Dir(Filename) = "" Then
						'get the names of the files to process
	   FileList = Application.GetOpenFilename("NMON Files(*.csv;*.nmon),*.csv;*.nmon", 1, "Select NMON file(s) to be processed", , True)
	   If VarType(FileList) <> vbBoolean Then Numfiles = UBound(FileList)
					       'write them to the sheet for sorting
	   If Numfiles = 0 Then Exit Sub
	   For i = 1 To Numfiles
	      MyCells.Offset(i, 0) = FileList(i)
	   Next
	Else
						'we have filelist - build a list of names
	   Open Filename For Input As #1
	   Do Until EOF(1)
	      Line Input #1, Fname
	      If Fname = "" Then Exit Do
	      sTemp = Dir(Fname)
	      fPath = Left(Fname, InStrRev(Fname, "\"))
	      Do
		If sTemp = "" Then Exit Do
		Numfiles = Numfiles + 1
		MyCells.Offset(Numfiles, 0) = fPath & sTemp
		sTemp = Dir
	      Loop
	   Loop
	   Close (1)
	   If Numfiles = 0 Then
	      MsgBox ("No valid files in FileList")
	      Exit Sub
	   End If
	End If
						'sort the names into ascending sequence
	MyCells.Resize(Numfiles + 1, 1).Sort Key1:=MyCells, Header:=xlYes
						'and store them in the Filelist array
	ReDim FileList(Numfiles)
	For i = 1 To Numfiles
	   FileList(i) = MyCells(i + 1, 1).Value
	   MyCells(i + 1, 1).Clear
	Next

End Sub
                                       'last mod v3.3.h
Sub Main(code As Integer)
	Dim aa0 As String
	'Public Batch as Integer
	'Public Delim As String                 'Delimiter used in .csv file
	Dim FileList As Variant
	'Public FileName As String              'Qualified name of the input file
	Dim i As Integer                        'counter for main loop
	Dim MyCells As Range                    'temp var
	Dim n As Integer                        'temp var
	Dim Numfiles As Integer                 'Number of files to process
	Dim NmonFile As Workbook                'Pointer to the output file
	Dim Output_Filename As String           'Qualified name of the output file
	Dim PID As Double                       'Process Id of spawned SORT process
	Dim PHn As Double                       'Process Header of spawned SORT process
	Dim RawData As Workbook                 'Pointer to the input file
	Dim Sortfile As String                  'name of temp file created by SORT process
	Dim splitpos As Integer                 'Pointer to last "." in input filename
	Dim SummaryFile As Workbook
	Dim Sheet1 As Worksheet
	Dim fs As Variant
	
	Batch = code
	Set fs = CreateObject("Scripting.FileSystemObject")
	Call GetSettings(FileList, Numfiles)
	If Numfiles = 0 Then
	   Batch = 0
	   Exit Sub
	End If
	    
	Application.ScreenUpdating = False      'set to True for debugging/False to improve performance
	
	Dim fileExtension As String
	If CInt(Val(Application.Version)) >= 12 Then
	    MaxRows = 1048576
	    fileExtension = ".xlsx"
	Else
	    MaxRows = 65536
	    fileExtension = ".xls"
	End If
	
	If Batch = 0 Then UserForm1.Show False
	
	If Merge <> "NO" Then Call MergeFiles(FileList, Numfiles)
	If Merge = "ONLY" Or Numfiles = 0 Then
	   UserForm1.Hide
	   Application.ScreenUpdating = True
	   Batch = 0
	   Exit Sub
	End If
	
	If Numfiles > 1 And Batch = 0 Then
					       'create a summary file
	   Set SummaryFile = Workbooks.Add(xlWBATWorksheet)
	   Set Sheet1 = SummaryFile.Worksheets(1)
	   Sheet1.Range("A1") = "Hostname"
	   Sheet1.Range("B1") = "Snapshots"
	   Sheet1.Range("C1") = "Start"
	   Sheet1.Range("D1") = "Filename"
	   Sheet1.Range("E1") = "Errors"
	   Sheet1.Rows(1).Font.Bold = True
	End If
	
	For i = 1 To Numfiles
						'initialisation
	   ErrMsg = ""
	   Filename = FileList(i)
	   Snapshots = 0
	   Sortfile = ""
	   If FirstTime > 0 Then Call GetIntervals(Filename)
	   With UserForm1
	      .Caption = Right(Filename, 40) & " (" & CStr(i) & " of " & CStr(Numfiles) & ")"
	      .Label1.Caption = "Opening file"
	      .Repaint
	   End With
						'construct the output filename
	   Output_Filename = Left(Filename, InStrRev(Filename, ".")) & "nmon" & fileExtension
	   If OutDir <> "" Then
	      splitpos = Len(Output_Filename) - InStrRev(Output_Filename, "\")
	      Output_Filename = OutDir & Right(Output_Filename, splitpos)
	   End If
						'bypass/delete file if it exists
	   If Dir(Output_Filename) <> "" Then
	      If Numfiles > 1 And ReProc = False Then
		 ErrMsg = "File has already been processed and REPROC=NO"
		 GoTo EndLoop
	      End If
	      fs.deletefile filespec:=Output_Filename
	   End If
						'open file
	   Set RawData = Workbooks.Open(Filename:=Filename, Format:=5)
						'if we have >MaxRows lines
	   If IsEmpty(RawData.Worksheets(1).Cells(MaxRows, 1)) Then
						'sort the input file if required
	      If SortInp Then
		 UserForm1.Label1.Caption = "Sorting file"
		 UserForm1.Repaint
		 RawData.Worksheets(1).Columns("A:A").Sort _
		    Key1:=RawData.Worksheets(1).Range("A1"), Order1:=xlAscending
	      End If
	   Else
						'check to see if the file has been sorted
	      Set MyCells = RawData.Worksheets(1).UsedRange.Find("CPU0", LookAt:=xlPart)
	      If (Left(MyCells(2, 1), 6) <> Left(MyCells(1, 1), 6)) Then
		 If SortInp Then
		    UserForm1.Label1.Caption = "Sorting file"
		    UserForm1.Repaint
						'otherwise we have to fork a process
		    Sortfile = Left(Filename, InStrRev(Filename, ".")) & "csv"
		    RawData.Close
		    aa0 = "SORT /l C /rec 8192 """ & Filename & """ /o """ & Sortfile & """"
		    PID = Shell(aa0, vbMinimizedNoFocus)
		    PHn = OpenProcess(SYNCHRONIZE, True, PID)
		    PID = WaitForSingleObject(PHn, INFINITE)
		    PID = CloseHandle(PHn)
		    If Dir(Sortfile) = "" Then
		       ErrMsg = "SORT command failed"
		       GoTo EndLoop
		    End If
		    Filename = Sortfile
		    Set RawData = Workbooks.Open(Filename:=Filename, Format:=5)
		 Else
		    MsgBox ("Input file needs to be sorted (>MaxRows lines AND SortInp=NO) " & Filename)
		    GoTo EndLoop
		 End If
	      End If
	   End If
						'avoid bizarre Excel bug if 1st line blank
	      Set MyCells = RawData.Worksheets(1).Columns("A:A")
	      If MyCells.Cells(1, 1) = "" Then Set MyCells = RawData.Worksheets(1).Range("A2:A" & CStr(MaxRows))
						'and parse it
	      UserForm1.Label1.Caption = "Parsing file"
	      UserForm1.Repaint
	      MyCells.TextToColumns Destination:=MyCells.Cells(1, 1), _
		 DataType:=xlDelimited, Other:=True, OtherChar:=Delim, _
		 DecimalSeparator:=DecSep, ThousandsSeparator:=ThouSep
	   
	   Set NmonFile = Workbooks.Add
	   Application.Calculation = xlCalculationAutomatic
	       
	   CPUrows = 0
	   dLPAR = False
	   NumCPUs = 0
	   NumDisk = 1
	   Call NMON(RawData, NmonFile)
	   
	   If NumCPUs > 0 Then
	      Application.DisplayAlerts = True
	      NmonFile.Sheets(1).Activate
		   
	      UserForm1.Label1.Caption = "Saving file"
	      If Numfiles > 1 Or Batch = 1 Then
		 NmonFile.SaveAs Filename:=Output_Filename
		 If Numfiles > 1 Or Batch = 1 Then NmonFile.Close
	      Else
		 Application.Dialogs(xlDialogSaveAs).Show Output_Filename
	      End If
	   Else
	      NmonFile.Close savechanges:=False
	      ErrMsg = "No valid input data! NMON run may have failed."
	   End If
	EndLoop:
					       'delete the sortfile if needed
	   If (Sortfile <> "" And Dir(Sortfile) <> "") Then fs.deletefile filespec:=Sortfile
					       'delete the merged file if needed
	   If Merge = "YES" Then fs.deletefile filespec:=FileList(1)
	
	   If Numfiles > 1 And Batch = 0 Then
					       'update summary file
	      If Snapshots > 0 Then
		 Sheet1.Cells(i + 1, 1) = Host
		 Sheet1.Columns(1).AutoFit
		 Sheet1.Cells(i + 1, 2) = Snapshots
		 Sheet1.Columns(2).AutoFit
		 Sheet1.Cells(i + 1, 3) = Start
		 Sheet1.Cells(i + 1, 3).NumberFormat = "dd-mmm-yy hh:mm:ss"
		 Sheet1.Columns(3).AutoFit
	      End If
	      If Dir(Output_Filename) <> "" Then
		 Sheet1.Hyperlinks.Add Anchor:=Sheet1.Range("D1").Offset(i, 0), _
		    Address:=Output_Filename, TextToDisplay:=Output_Filename
	      End If
	      Sheet1.Columns(4).ColumnWidth = 29
	      Sheet1.Cells(i + 1, 4).HorizontalAlignment = xlRight
	      Sheet1.Cells(i + 1, 5) = ErrMsg
	   End If
	   If Batch = 0 And Numfiles = 1 And ErrMsg <> "" Then MsgBox (ErrMsg)
	Next i
	
	UserForm1.Hide
	Application.ScreenUpdating = True
	If (Batch = 0 And Not (Numfiles = 1 And ErrMsg <> "")) Then ThisWorkbook.Close False
	Exit Sub

End Sub
                                       'last mod v3.4a
Sub MergeFiles(FileList As Variant, Numfiles As Integer)
	Dim aa0 As String, aa1 As String
	Dim fnin As Variant, fnout As Integer
	Dim Filename As String                 'output filename (ddmmyyhhmm.nmon)
	Dim fs As Variant                      'Filesystem object
	Dim i As Integer, j As Integer         'loop counters
	Dim Tlen As Integer                    'length of timestamp
	Dim Tn As Integer                      'current timestamp
	Dim hostname As String                 'hostname of first file"
	Dim t1 As String                       'time of first file
	Dim d1 As String                       'date of first file
	Dim t2 As String                       'time of first file
	Dim d2 As String                       'date of first file
	
	With UserForm1
	   .Caption = "Merging files"
	   .Repaint
	End With
					       'open a temp file
	Filename = CurDir & "\" & "merged_" & Format(Now(), "yyyymmdd_hhmm") & ".nmon"
	fnout = FreeFile()
	Open Filename For Output As fnout
					       'copy the first file
	With UserForm1
	   .Label1.Caption = Right(FileList(1), 25)
	   .Show False
	   .Repaint
	End With
	
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set fnin = fs.OpenTextFile(FileList(1), 1, 0)
						  'extract only the sections we need
	Do While fnin.AtEndOfStream <> True
	   aa0 = fnin.readline
	   aa1 = Left(aa0, 4)
	   If aa1 = "ZZZZ" Then
	      Tlen = InStr(7, aa0, Delim) - 7
	      Tn = Mid(aa0, 7, Tlen)
	   End If
	   If Left(aa0, 9) = "AAA,host," Then hostname = Mid(aa0, 10)
	   If Left(aa0, 9) = "AAA,time," Then t1 = Mid(aa0, 10)
	   If Left(aa0, 9) = "AAA,date," Then d1 = Mid(aa0, 10)
	   If Not ((aa1 = "TOP," Or aa1 = "UARG") And NoTop) Then
	      If Left(aa0, 14) <> "AAA,snapshots," Then
		 j = InStr(1, aa0, ",T")
		 If j > 0 Then
		    aa0 = Left(aa0, j + 1) & Format(Tn, "000000") & Right(aa0, Len(aa0) - j - Tlen - 1)
		 End If
		 Print #fnout, aa0
	      End If
	   End If
	Loop
	fnin.Close
	
	For i = 2 To Numfiles
	   With UserForm1
	      .Label1.Caption = Right(FileList(i), 25)
	      .Show False
	      .Repaint
	   End With
	   Set fnin = fs.OpenTextFile(FileList(i), 1, 0)
					       'skip headers (assumes file is unsorted)
	   Do
	      aa0 = fnin.readline
	      If Left(aa0, 4) = "ZZZZ" Then Exit Do
	   Loop
	   Do While fnin.AtEndOfStream <> True
	      aa1 = Left(aa0, 4)
	      If aa1 = "ZZZZ" Then
		 Tlen = InStr(7, aa0, Delim) - 7
		 Tn = Tn + 1
		 t2 = Mid(aa0, Tlen + 8, 8)
		 d2 = Right(aa0, 11)
	      End If
	      If Not ((aa1 = "TOP," Or aa1 = "UARG") And NoTop) Then
		 j = InStr(1, aa0, ",T")
		 If j > 0 Then
		    aa0 = Left(aa0, j + 1) & Format(Tn, "000000") & Right(aa0, Len(aa0) - j - Tlen - 1)
		    Print #fnout, aa0
		 End If
	      End If
	      aa0 = fnin.readline
	   Loop
	   fnin.Close
	Next i
	finish:
	Print #fnout, "AAA,snapshots," & CStr(Tn)
	Close fnout
						'rename the file
	d1 = Format(d1, "yymmdd")
	t1 = Format(t1, "hhmm")
	d2 = Format(d2, "yymmdd")
	t2 = Format(t2, "hhmm")
	aa0 = CurDir & "\" & hostname & "_" & d1 & "_" & t1 & "_to_" & d2 & "_" & t2 & ".nmon"
	Name Filename As aa0
	FileList(1) = aa0
	Numfiles = 1
End Sub
                                      'last mod v3.3.g2
Sub NMON(Book1 As Workbook, NmonFile As Workbook)
	'Public Esoteric As String              'Either "hdiskpower" or "dac"
	'Public GotEMC As Variant               'True if either EMC or FAStT present
	'Public GotESS As Variant               'True if EMC/ESS or FAStT present
	'Public NumCPUs As Integer              'number of CPU sections
	Dim CPUList(1 To 1024) As String        'names of CPUnn sheets
	Dim CurrentRow As Long                  'pointer to end of range
	Dim Elapsed As Single                   'elapsed time in seconds
	Dim FirstSheets() As String             'Array of names of sheets to be deleted
	Dim n As Long                           'temp var
	Dim numrows As Long                     'number or rows in current section
	Dim numcols As Integer                  'number of columns in current section
	Dim Section As Range                    'Range for cut/paste
	Dim RawData As Worksheet
	Dim SectionName As String               'name of current section
	Dim Sheet1 As Worksheet                 'new sheet
	Dim StartRow As Long                    'pointer to start of range
	Dim StrayLines As Integer               'counter for stray lines
	Dim StraySheet As Worksheet             'pointer to StrayLines sheet
	
	Set RawData = Book1.Worksheets(1)
	Elapsed = Timer()
	With UserForm1
	   .Label1.Caption = "Starting Analysis"
	   .Repaint
	End With
	Call DefineStyles
						'build a list of sheet names for later deletion
	ReDim FirstSheets(Application.SheetsInNewWorkbook)
	For n = 1 To Application.SheetsInNewWorkbook
	   FirstSheets(n) = Worksheets(n).Name
	Next n
	'following section needs revising - move to PP_DISK to avoid problems with >65K files
	'and also to handle mixed FAStT/EMC/ESS configs properly
						'see if we have an EMC system
	Esoteric = "hdiskpower"
	Set Section = RawData.UsedRange.Find(Esoteric, LookAt:=xlPart)
	GotEMC = Not (Section Is Nothing)
	If GotEMC = False Then
	    Esoteric = "emcpower"
	    Set Section = RawData.UsedRange.Find(Esoteric, LookAt:=xlPart)
	    GotEMC = Not (Section Is Nothing)
	End If
						'see if we have a FAStT subsystem
	If GotEMC = False Then
	   Esoteric = "dac0"
	   Set Section = RawData.UsedRange.Find(Esoteric, LookAt:=xlWhole)
	   GotEMC = Not (Section Is Nothing)
	   Esoteric = "dac"
	End If
	GotESS = GotEMC
	
	StartRow = 1
	CurrentRow = 1
	SMTmode = 1
	If RawData.Range("A1") = "" Then CurrentRow = 2   'horrible fix for Excel bug
	Do 'Until SectionName = ""
	   Call GetNextSection(RawData, SectionName, StartRow, CurrentRow)
	   If SectionName = "" Then
	      Exit Do
	   ElseIf SectionName = "TOP" Then
	      StartRow = StartRow + 1
	      RawData.Cells(StartRow, 2).Value = "PID"
	   ElseIf SectionName = "UARG" Then
	      RawData.Cells(StartRow, 2).Value = "Time"
	   ElseIf CPUmax > 0 And InStr(Left$(SectionName, 4), "CPU") > 0 Then   'skip CPU sheets
	      If Left(SectionName, 3) = "CPU" Then
		 If Val(Mid(SectionName, 4)) > CPUmax Then
		    NumCPUskipped = NumCPUskipped + 1
		    GoTo EndSect
		 End If
	      Else
		 If Val(Mid(SectionName, 5)) > CPUmax Then GoTo EndSect
	      End If
	   End If
	   numrows = CurrentRow - StartRow
	   If numrows > MaxRows - 256 Then numrows = MaxRows - 256  'leave space for totals etc.
	   n = StartRow + numrows - 1
	   Set Section = RawData.Range("B" & CStr(StartRow) & "..IV" & CStr(n))
						  'if a valid section, build a sheet
	   If numrows > 1 And SectionName <> "" And SectionName <> "ERROR" Then
	      UserForm1.Label1.Caption = "Analysing: " & SectionName
	      UserForm1.Repaint
	      Sheets.Add.Name = SectionName
	      Set Sheet1 = Worksheets(SectionName)
	      Sheet1.Move After:=Sheets(Sheets.Count)
	      Section.Copy Sheet1.Range("A1")
	      Application.CutCopyMode = False
	      Call GetLastColumn(Sheet1)
						  'Do any Post Processing
	      If Left$(SectionName, 3) = "AAA" Then
		 Call PP_AAA(numrows, Sheet1)
	      ElseIf Left$(SectionName, 3) = "BBB" Then
		 Call PP_BBB(numrows, Sheet1)
	      ElseIf Left$(SectionName, 4) = "CPU_" Then
		 Call PP_CPU(numrows, Sheet1)
		 If CPUrows < 3 Then
		    ErrMsg = "Only one interval - processing terminated"
		    Snapshots = 1
		    NumCPUs = NumCPUs + 1         'avoid "no valid data" message
		    Book1.Close savechanges:=False
		    GoTo FinishUp
		 End If
	      ElseIf InStr(Left$(SectionName, 4), "CPU") > 0 Then
		 If Left(SectionName, 3) = "CPU" Then
		    If CPUmax > 0 And Val(Mid(SectionName, 4)) > CPUmax Then GoTo EndSect
		    NumCPUs = NumCPUs + 1
		    CPUList(NumCPUs) = SectionName
		 Else
		    If CPUmax > 0 And Val(Mid(SectionName, 5)) > CPUmax Then GoTo EndSect
		 End If
		 Call PP_CPU(numrows, Sheet1)
	      ElseIf Left$(SectionName, 2) = "DG" Then
		 GotESS = True
		 Call PP_DG(numrows, Sheet1)
	      ElseIf Left$(SectionName, 4) = "DISK" Then
		 Call PP_DISK(numrows, Sheet1)
	      ElseIf SectionName = "DONATE" Then
		 Call PP_DONATE(numrows, Sheet1)
	      ElseIf Left$(SectionName, 3) = "ESS" Then
		 Call PP_ESS(numrows, Sheet1)
	      ElseIf SectionName = "FILE" Then
		 Call PP_FILE(numrows, Sheet1)
	      ElseIf SectionName = "FRCA" Then
		 Call PP_FRCA(numrows, Sheet1)
	      ElseIf SectionName = "IOADAPT" Then
		 Call PP_IOADAPT(numrows, Sheet1)
	      ElseIf SectionName = "IP" Then
		 Call PP_IP(numrows, Sheet1)
	      ElseIf Left$(SectionName, 3) = "JFS" Then
		 Call PP_JFS(numrows, Sheet1)
	      ElseIf SectionName = "LAN" Then
		 Call PP_LAN(numrows, Sheet1)
	      ElseIf SectionName = "LARGEPAGE" Then
		 Call PP_LPAGE(numrows, Sheet1)
	      ElseIf SectionName = "LPAR" Then
		 Call PP_LPAR(numrows, Sheet1)
	      ElseIf SectionName = "MEM" Then
		 Call PP_MEM(numrows, Sheet1)
	      ElseIf SectionName = "MEMAMS" Then
		 Call PP_MEMAMS(numrows, Sheet1)
	      ElseIf SectionName = "MEMNEW" Then
		 Call PP_MEMNEW(numrows, Sheet1)
	      ElseIf Left(SectionName, 8) = "MEMPAGES" Then
		 Call PP_MEMPAGES(numrows, Sheet1)
	      ElseIf SectionName = "MEMREAL" Then
		 Call PP_MEMREAL(numrows, Sheet1)
	      ElseIf SectionName = "MEMUSE" Then
		 Call PP_MEMUSE(numrows, Sheet1)
	      ElseIf SectionName = "MEMVIRT" Then
		 Call PP_MEMVIRT(numrows, Sheet1)
	      ElseIf SectionName = "NET" Then
		 Call PP_NET(numrows, Sheet1)
	      ElseIf SectionName = "PAGE" Then
		 Call PP_PAGE(numrows, Sheet1)
	      ElseIf SectionName = "PAGING" Then
		 Call PP_PAGING(numrows, Sheet1)
	      ElseIf SectionName = "POOLS" Then
		 Call PP_POOLS(numrows, Sheet1)
	      ElseIf SectionName = "PROC" Then
		 Call PP_PROC(numrows, Sheet1)
	      ElseIf SectionName = "PROCAIO" Then
		 Call PP_PROCAIO(numrows, Sheet1)
	      ElseIf Left$(SectionName, 3) = "RAW" Then
		 Call PP_RAW(numrows, Sheet1)
	      ElseIf SectionName = "SUMMARY" Then
		 Call PP_SUMMARY(numrows, Sheet1)
	      ElseIf SectionName = "TOP" Then
		 Call PP_TOP(numrows, Sheet1)
	      ElseIf SectionName = "VM" Then
		 Call PP_VM(numrows, Sheet1)
	      ElseIf SectionName = "UARG" Then
		 Call PP_UARG(numrows, Sheet1)
	      ElseIf Left$(SectionName, 3) = "WLM" Then
		 Call PP_WLM(numrows, Sheet1)
	      ElseIf SectionName = "ZZZZ" Then
		 Call PP_ZZZZ(numrows, Sheet1)
		 Exit Do   'ignore anything after the ZZZZ section
	      Else
		 Call PP_DEFAULT(numrows, Sheet1)
	      End If
	   Else
	      If SectionName = "ERROR" Then   'ERROR section doesn't have a header
		 Sheets.Add.Name = SectionName
		 Set Sheet1 = Worksheets(SectionName)
		 Sheet1.Range("A1") = "Errors reported by NMON in nmon file - no interval filter"
		 Sheet1.Cells(2, 1).Resize(numrows, 255).Value = Section.Value
		 Application.CutCopyMode = False
	      Else
		 If StrayLines = 0 Then
		    ErrMsg = "Some lines discarded"
		    Sheets.Add.Name = "StrayLines"
		    Set StraySheet = Worksheets("StrayLines")
		    StraySheet.Range("B1") = "Following lines discarded after parsing"
		 End If
		 If SectionName <> "" And StrayLines < 50 Then
		    StrayLines = StrayLines + 1
		    StraySheet.Cells(StrayLines + 1, 1) = SectionName
		    StraySheet.Cells(StrayLines + 1, 2).Resize(numrows, 255).Value = Section.Value
		    Application.CutCopyMode = False
		 End If
	      End If
	   End If
	EndSect:
	Loop Until SectionName = ""
	Book1.Close savechanges:=False
	If NumCPUs = 0 Then Exit Sub
					       'produce the CPU_SUMM and SYS_SUMM sheets
	Call CPU_SUMM(CPUList())
	Call SYS_SUMM
					       'finish up
	FinishUp:
	Call TidyUp(CPUList(), FirstSheets())
					       'Convert, print/publish the charts if necessary
	If Output <> "CHARTS" Then Call OutputPICS
	If Pivot Then Call CreatePivot
	If Worksheets(1).Name <> "SYS_SUMM" Then
	   Worksheets(1).Move After:=Sheets("ZZZZ")
	End If
					       'finish up
	If SheetExists("AAA") Then
	   Elapsed = Timer() - Elapsed
	   Set Section = Sheets("AAA").Range("A1").End(xlDown)
	   Section.Offset(1, 0) = "elapsed"
	   Section.Offset(1, 1) = Format(Elapsed, "#.00") & " seconds"
	   UserForm1.Label1.Caption = "Analysis Complete (" & Elapsed & " seconds)"
	   UserForm1.Repaint
		
	   If (StrayLines > 0) And (NumCPUs > 0) Then
	      Sheets("StrayLines").Move After:=Sheets("AAA")
	   End If
	End If
End Sub
                                       'last mod v3.2.0
Sub OutputPICS()
	Dim temp As String
	Dim Chart1 As ChartObject
	Dim i As Long, j As Long, k As Long
	Dim MyCharts As Worksheet
	Dim myRange As Range
	Dim nc As Integer                      'number of charts on Charts page
	Dim Output_Filename As String          'name for HTML output
	Dim pageDepth As Integer               'number of rows that can fit on a printed page
	Dim PicRows(400) As Long               'location of bitmaps for setting page breaks
	Dim myVar As Variant                   'temp var
	Dim splitpos As String
	'Public Summary As String
	Dim Sheet1 As Worksheet
	
	UserForm1.Label1.Caption = "Converting charts to pictures"
	UserForm1.Repaint
					       'Create a chart sheet
	Sheets.Add.Name = "Charts"
	Set MyCharts = Worksheets("Charts")
	MyCharts.Move Before:=Worksheets(1)
	If LScape Then MyCharts.PageSetup.Orientation = xlLandscape
					       'now move each chart to the charts sheet
	i = 1: nc = 0: myVar = True
	For Each Sheet1 In ActiveWorkbook.Sheets
	   For Each Chart1 In Sheet1.ChartObjects
	      nc = nc + 1
	      Chart1.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture, Size:=xlScreen
	      MyCharts.Paste Destination:=MyCharts.Cells(i, 1)
	      i = i + Chart1.BottomRightCell.Row - Chart1.TopLeftCell.Row + 1
	      PicRows(nc) = i
	      Chart1.Delete
	    Next
	Next
	If Output = "PICTURES" Then
	   Exit Sub
	ElseIf Output = "PRINT" Then
					       'PRINT option - adjust horizontal page breaks
	   With ActiveSheet.PageSetup
	      .CenterHeader = "&""Arial,Bold""&14" & Host & " " & RunDate
	      .CenterHorizontally = True
	      .CenterVertically = True
	      .LeftFooter = "NMON " & progname
	      .RightFooter = "Analyser " & Analyser
	   End With
	   If nc > 1 Then
						  'make sure that a chart doesn't span pages
	      ActiveWindow.View = xlPageBreakPreview
	      temp = MyCharts.HPageBreaks(1).Location.Address(True, True, xlR1C1)
	      pageDepth = Mid(temp, 2, InStr(1, temp, "C") - 2) - 1
	      MyCharts.Cells(nc * pageDepth, 1) = "break"
	      UserForm1.Label1.Caption = "Adjusting horizontal page breaks"
	      UserForm1.Repaint
						  'by dynamically adjusting the pagebreaks
	      i = pageDepth: k = 1
	      For j = 2 To nc
		 If PicRows(j) >= i Then
		    Set MyCharts.HPageBreaks(k).Location = Cells(PicRows(j - 1), 1)
		    k = k + 1
		    i = PicRows(j - 1) + pageDepth
		 End If
	      Next
	      MyCharts.Cells(nc * pageDepth, 1).Clear
	   End If
	   If Printer = "PREVIEW" Then
	      UserForm1.Hide
	      ActiveSheet.PrintPreview
	      UserForm1.Show
	   Else
	      ActiveSheet.PrintOut Copies:=Copies, ActivePrinter:=Printer
	   End If
	   ActiveWindow.View = xlNormalView
						  'publish pictures to the web
	ElseIf Output = "WEB" Then
	   UserForm1.Label1.Caption = "Generating HTML"
	   UserForm1.Repaint
					       'construct the output filename
	   Output_Filename = Left(Filename, InStrRev(Filename, ".")) & "nmon.htm"
	   If WebDir <> "" Then
	      splitpos = Len(Output_Filename) - InStrRev(Output_Filename, "\")
	      Output_Filename = WebDir & Right(Output_Filename, splitpos)
	   End If
					       'and generate html/bitmaps
	   ActiveWorkbook.WebOptions.AllowPNG = PNG
	   ActiveWorkbook.WebOptions.OrganizeInFolder = SubDir
	   With ActiveWorkbook.PublishObjects.Add(SourceType:=xlSourceSheet, _
	      Filename:=Output_Filename, Sheet:="Charts", Source:="", _
	      HtmlType:=xlHtmlStatic, Title:=Host & " " & RunDate)
	      .Publish (True)
	'     .AutoRepublish = False   (not supported on Excel 2000)
	   End With
	End If

End Sub
                                       'last mod v3.3.g2
Sub PP_AAA(numrows As Long, Sheet1 As Worksheet)
	'Public Host As String                  Hostname from AAA sheet
	'Public Linux as Variant                True if Linux
	'Public progname As String              Version of NMON/topas
	'Public RunDate As String               NMON run date from AAA sheet
	Dim i As Long
	Dim Labels As Range
	Dim temp As String
	Dim TimeAAA As Variant
					      'Delete any notes
	Set Labels = Sheet1.Range("A1:B" & CStr(numrows))
	For i = numrows To 1 Step -1
	   If Left(Labels.Cells(i, 1), 4) = "note" Then
	   Labels.Rows(i).EntireRow.Delete
	   numrows = numrows - 1
	   End If
	Next i
	
	Sheet1.Columns(1).Font.Bold = True
	Sheet1.Columns(1).AutoFit
	Sheet1.Columns(2).ColumnWidth = 10
	Sheet1.Columns(3).AutoFit
	Sheet1.Columns(4).AutoFit
	Labels.CreateNames Top:=False, Left:=True, bottom:=False, Right:=False
	Sheet1.Range("Date").NumberFormat = "dd-mmm-yy"
	Sheet1.Range("time").NumberFormat = "hh:mm:ss"
						'start time for batch summary
	Start = DateValue(Sheet1.Range("date"))
	TimeAAA = Sheet1.Range("time")
	If Not IsNumeric(TimeAAA) Then TimeAAA = TimeValue(TimeAAA)
	Start = Start + TimeAAA
	
	ActiveWorkbook.Names("time").Delete     'strange Excel bug fix
	Host = Sheet1.Range("Host").Value
	
	progname = Sheet1.Range("progname").Value
	RunDate = Sheet1.Range("Date").Value
	Snapshots = Sheet1.Range("snapshots").Value
					       'adjust first/last values if required for this file
	'If Last > Snapshots Then Last = Snapshots
	If First >= Snapshots Or Last < First + 1 Or First = 0 Then
	   If Batch = 0 Then MsgBox ("Invalid values for FIRST/LAST - values reset to 1/999999")
	   First = 1
	   Last = 999999
	End If
					       'record settings in output file
	Labels(numrows + 1, 1) = "analyser"
	Labels(numrows + 1, 2) = Analyser
	Labels(numrows + 2, 1) = "environment"
	Labels(numrows + 2, 2) = "Excel " & Application.Version & " on " & Application.OperatingSystem
	temp = "BATCH=" & CStr(Batch)
	temp = temp + ",FIRST=" & CStr(First) & ",LAST=" & CStr(Last)
	temp = temp + ",GRAPHS=" & Graphs
	temp = temp + ",OUTPUT=" & Output
	temp = temp + ",CPUmax=" & CStr(CPUmax)
	temp = temp + ",MERGE=" & Merge
	temp = temp + ",NOTOP=" & NoTop
	temp = temp + ",PIVOT=" & Pivot
	temp = temp + ",REORDER=" & Reorder
	temp = temp + ",TOPDISKS=" & CStr(TopDisks)
	Labels(numrows + 3, 1) = "parms"
	Labels(numrows + 3, 2) = temp
	temp = "GWIDTH = " & CStr(cWidth)
	temp = temp + ",GHEIGHT=" & CStr(cHeight)
	temp = temp + ",LSCAPE=" & LScape
	temp = temp + ",REPROC=" & ReProc
	temp = temp + ",SORTINP=" & SortInp
	Labels(numrows + 4, 1) = "settings"
	Labels(numrows + 4, 2) = temp
	Sheet1.Range("B1:B" & CStr(numrows)).HorizontalAlignment = xlLeft
						'see if we have a Linux system
	Set Labels = Sheet1.UsedRange.Find("Linux", LookAt:=xlWhole)
	Linux = Not (Labels Is Nothing)
						'check if this is topas
	Set Labels = Sheet1.UsedRange.Find("PTX:", LookAt:=xlPart)
	topas = Not (Labels Is Nothing)
	      
End Sub
                                       'v3.3.g2
Sub PP_BBB(numrows As Long, Sheet1 As Worksheet)
	Dim aa0 As String
	Dim i As Long
	Dim MyCells As Range
	Dim temp As Boolean
						'sort the data
	Sheet1.UsedRange.Sort Key1:=Sheet1.Range("A1"), Order1:=xlAscending, _
	   Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
	Sheet1.Columns("A").EntireColumn.Delete
	
	Select Case Sheet1.Name
	 Case "BBBB", "BBBL", "BBBN", "BBBVG"
	   Sheet1.Columns(1).AutoFit
	 Case "BBBC"
						'change to fixed pitch font
	   Set MyCells = Sheet1.Range("A1:A" & CStr(numrows))
	   MyCells.Font.Name = BBBFont
						'change fonts for hdisk/paging lines
	   For i = numrows To 1 Step -1
	      aa0 = MyCells.Item(i, 1).Value
	      If Left(aa0, 5) = "hdisk" Then
		 MyCells.Item(i, 1).Font.Bold = True
	      End If
	   Next i
	     
	 Case "BBBD"
	   Sheet1.Columns("A").ColumnWidth = 14.5
	   Sheet1.Columns("D").ColumnWidth = 17
	   Sheet1.Rows(2).Font.Bold = True
	   Set MyCells = Sheet1.Range("B3:B" & CStr(numrows))
	   For i = 1 To numrows
	       MyCells(i, 1) = MyCells(i, 1).Value
	   Next i
	 Case "BBBE"
	   Sheet1.Rows(2).Font.Bold = True
	   GotESS = True
						'define the lookup table for ESS analysis
	   If Sheet1.Range("B2") = "Name" Then
		 ActiveWorkbook.Names.Add Name:="VPATHS", RefersTo:="=BBBE!B3:IV" & CStr(numrows)
	      Else
		 ActiveWorkbook.Names.Add Name:="VPATHS", RefersTo:="=BBBE!C3:IV" & CStr(numrows)
	      End If
	 Case "BBBF"
	   Sheet1.Columns(3).AutoFit
	 Case "BBBP"
	   Sheet1.Columns(1).AutoFit
						'change to fixed pitch font
	   Set MyCells = Sheet1.Columns("B:C")
	   MyCells.Font.Name = BBBFont
						'check for SMT
	   Set MyCells = Sheet1.UsedRange.Find("PowerPC_Power7", LookAt:=xlPart)
	   SMTon = Not (MyCells Is Nothing)
	   If SMTon Then
	      Set MyCells = Sheet1.UsedRange.Find("-SMT-4", LookAt:=xlPart)
	      SMTon = Not (MyCells Is Nothing)
	      If SMTon Then
		 SMTmode = 4
	      Else
		 Set MyCells = Sheet1.UsedRange.Find("-SMT", LookAt:=xlPart)
		 SMTon = Not (MyCells Is Nothing)
		 If SMTon Then SMTmode = 2
	      End If
	   Else
	      Set MyCells = Sheet1.UsedRange.Find("-SMT", LookAt:=xlPart)
	      SMTon = Not (MyCells Is Nothing)
	      If SMTon Then SMTmode = 2
	   End If
	   Set MyCells = Sheet1.UsedRange.Find("Shared-", LookAt:=xlPart)
	    If Not (MyCells Is Nothing) Then
	       If CPUmax = 0 Then CPUmax = SMTmode 'set CPUmax for Shared LPARs
	    End If
	 Case "BBBR"
	   Sheet1.Columns(1).NumberFormat = "h:mm:ss"
	 Case "BBBV"
						'change to fixed pitch font
	   Set MyCells = Sheet1.Range("A1:A" & CStr(numrows))
	   MyCells.Font.Name = BBBFont
	End Select
		    
End Sub
                                       'last mod v3.3.g1
Sub PP_CPU(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject               'New Chart object
	Dim MyCells As Range                    'temp var
	'Public CPUs as Integer                 'number of CPUs as on AAA sheet
	'Public CPUrows As Long                 'number of rows on the CPU01 sheet
	'Public dLPAR As Variant                'True if dynamic LPAR has been used for CPUs
	'Public NumCPUs                         'current number of CPUnn sheets
	'Public Host As String                  'Hostname from AAA sheet
	'Public RunDate As String               'NMON run date from AAA sheet
	'Public T1 As String                    'Time stamp for first interval
	
						'quick fix to missing header problem
	If Sheet1.Name = "CPU_ALL" And Sheet1.Range("B1") <> "User%" Then
	   Sheet1.Rows(1).Insert
	   Sheet1.Range("A1").Value = "CPU Total " & Worksheets("AAA").Range("Host").Value
	   Sheet1.Range("B1").Value = "User%"
	   Sheet1.Range("C1").Value = "Sys%"
	   Sheet1.Range("D1").Value = "Wait%"
	   Sheet1.Range("E1").Value = "Idle%"
	   Sheet1.Range("G1").Value = "CPUs"
	   numrows = numrows + 1
	End If
						'save number of intervals for later
	If CPUrows = 0 Then
	   t1 = Sheet1.Range("A2")
						'bodge for Linux
	   If (Linux And t1 = "T0002") Then
	      t1 = "T0001"
	      Sheet1.Rows(2).Insert
	      Sheet1.Range("A2").Value = t1
	      numrows = numrows + 1
	   End If
	   CPUrows = numrows
	End If
						'handle missing intervals for dLPAR
	If numrows < CPUrows Then
	   Sheet1.Range("A2", "G" & CStr(CPUrows - numrows + 1)).EntireRow.Insert
	   Sheet1.Range("A2").Value = t1
	   numrows = CPUrows
	   dLPAR = True
	End If
	Call DelInt(Sheet1, numrows, 1)
						'Calculate CPU totals for graphs
	Sheet1.Range("F1").Value = "CPU%"
	Sheet1.Range("F2").Value = "=IF(B2="""","""",B2+C2)"
	Set MyCells = Sheet1.Range("F2:F" & CStr(numrows))
	MyCells.FillDown
	MyCells.Value = MyCells.Value
	       
	Sheet1.Cells(numrows + 2, 1) = "Avg"
	Sheet1.Cells(numrows + 2, 2) = "=AVERAGE(B2:B" & CStr(numrows) & ")"
	Set MyCells = Sheet1.Range("B" & CStr(numrows + 2) & ":F" & CStr(numrows + 2))
	MyCells.FillRight
	MyCells.Value = MyCells.Value
	If Graphs = "LIST" And Not Checklist(Sheet1.Name) Then Exit Sub
						'Produce a graph of CPU components
	Set MyCells = Sheet1.Range("A1..D" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, cTop + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	CategoryLabels:=1, SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   If Left(Sheet1.Name, 3) = "CPU" Then .Axes(xlValue).MaximumScale = 100
	   Call ApplyStyle(Chart1, 2, 3)
	End With
	If Sheet1.Range("G2") = "" Then Exit Sub
						'produce a graph of CPU count
	If Sheet1.Name = "CPU_ALL" And SMTon Then Sheet1.Range("G1") = "Logical CPUs (SMTmode=" & CStr(SMTmode) & ")"
	Set MyCells = Sheet1.Range("G1:G" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	   SeriesLabels:=1, HasLegend:=True
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = Sheet1.Range(Cells(2, 1), Cells(CPUrows, 1))
	   Call ApplyStyle(Chart1, 2, 1)
	   .HasTitle = False
	End With
End Sub
                                       'last mod v3.3.g2
Sub CPU_SUMM(CPUList() As String)
	Dim aa0 As String
	Dim Chart1 As ChartObject               'new chart object
	Dim CPUData As Range                    'data to be charted
	'Public CPUrows As Long                 'number of rows on the CPU01 sheet
	'Public dLPAR as Variant                'True if dLPAR for CPUs
	'Public Host As String                  'Hostname from AAA sheet
	Dim n As Integer                        'loop counter
	'Public NumCPUs As Integer              'number of CPU sections
	'Public RunDate As String               'NMON run date from AAA sheet
	Dim SectionName As String               'constant = "CPU_SUMM"
	Dim Sheet1 As Worksheet                 'pointer to CPU_SUMM sheet
						'Produce a sheet containing summary data
	SectionName = "CPU_SUMM"
	UserForm1.Label1.Caption = "Creating: " & SectionName
	UserForm1.Repaint
	
	Sheets.Add.Name = SectionName
	If SheetExists("CPU_ALL") Then Sheets(SectionName).Move After:=Worksheets("CPU_ALL")
	Set Sheet1 = Worksheets(SectionName)
	Sheet1.Range("A1").Value = SectionName
	Sheet1.Range("B1").Value = "User%"
	Sheet1.Range("C1").Value = "Sys%"
	Sheet1.Range("D1").Value = "Wait%"
	Sheet1.Range("E1").Value = "Idle%"
	
	CPUrows = Worksheets("AAA").Range("snapshots").Value + 1
	Application.Calculation = xlCalculationManual
	For n = 1 To NumCPUs
	   Sheet1.Cells(n + 1, 1) = "CPU" & Format(Mid(CPUList(n), 4), "0##")
	   Sheet1.Cells(n + 1, 2) = "=" & CPUList(n) & "!B" & CStr(CPUrows + 2)
	Next n
	Set CPUData = Sheet1.Range("B2:E" & CStr(NumCPUs + 1))
	CPUData.FillRight
	Application.Calculation = xlCalculationAutomatic
	Application.Calculate
	CPUData.Value = CPUData.Value
	CPUData.NumberFormat = "#0.0"
	Sheet1.Columns("A:IV").Sort Key1:=Sheet1.Range("A1"), Order1:=xlAscending, Header:=xlYes
	
	If Graphs = "LIST" And Not Checklist(Sheet1.Name) Then Exit Sub
						'and now produce a graph by CPU/Thread
	aa0 = " Processor "
	If SMTon Then aa0 = " Thread "
	Set CPUData = Sheet1.Range("A1:D" & CStr(NumCPUs + 1))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, cTop + NumCPUs * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=CPUData, PlotBy:=xlColumns, _
	   CategoryLabels:=1, SeriesLabels:=1, _
	   Title:="CPU by" & aa0 & Host & "  " & RunDate & "    (" & CStr(NumCPUskipped) & " threads not shown)"
						'apply customisation
	With Chart1.Chart
	     .Axes(xlValue).MaximumScale = 100
	     Call ApplyStyle(Chart1, 0, 3)
	End With
	
	Sheet1.Range("B2").Select
	ActiveWindow.FreezePanes = True
	Sheet1.Range("A1").Select
	ActiveWindow.ScrollRow = NumCPUs + 2
End Sub
                                       'last mod v3.3.A
Sub PP_DEFAULT(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject               'new chart object
	'Public ColNum as Integer
	Dim Graphdata As Range                  'range for charting
	'Public Host As String                  'Hostname from AAA sheet
	'Public LastColumn As String
	'Public RunDate As String               'NMON run date from AAA sheet
	Dim SectionName As String
	    
	SectionName = Sheet1.Name
	If Application.WorksheetFunction.Max(Sheet1.Range("B2:IV" & CStr(numrows))) = 0 Then
	   Application.DisplayAlerts = False
	   Sheet1.Delete
	   Application.DisplayAlerts = True
	   Exit Sub
	End If
	Call DelInt(Sheet1, numrows, 1)
	If numrows < 1 Then Exit Sub
	If Graphs = "LIST" And Not Checklist(Sheet1.Name) Then Exit Sub
						'produce avg/max graph
	Set Graphdata = avgmax(numrows, Sheet1, 0)
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=Graphdata, PlotBy:=xlRows, SeriesLabels:=1, _
	    Title:=Sheet1.Range("A1").Value & "  " & RunDate
	With Chart1.Chart                      'apply customisation
	   .SeriesCollection(1).XValues = Sheet1.Range(Cells(1, 2), Cells(1, ColNum))
	   Call ApplyStyle(Chart1, 0, 3)
	End With
						'produce line graph
	Set Graphdata = Sheet1.Range("A1:" & LastColumn & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + rH * numrows, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=Graphdata, Gallery:=xlLine, Format:=2, _
	    Title:=Sheet1.Range("A1").Value & "  " & RunDate, _
	    CategoryLabels:=1, SeriesLabels:=1, HasLegend:=True
						'apply customisation
	With Chart1.Chart
	   Call ApplyStyle(Chart1, 1, ColNum - 1)
	End With
	      
	End Sub
					       'last mod v3.3.0
	Sub PP_DG(numrows As Long, Sheet1 As Worksheet)
	    Dim SectionName As String
	    
	    SectionName = Sheet1.Name
	    Call DelInt(Sheet1, numrows, 1)
	    Call DiskGraphs(numrows, SectionName, True)
End Sub
                                       'v3.3.g1
Sub PP_DISK(numrows As Long, Sheet1 As Worksheet)
	Dim aa0 As String                   'temp var
	'Public bSize As Variant            'True if DISKBSIZE present
	Dim DelRange As Range               'Columns to be deleted
	'Public DiskSort As Variant         'Value of DISKSORT field
	'Public Esoteric As String          'Either "hdiskpower" or "dac"
	Dim fed As Integer                  'column number of first esoteric
	Dim found As Variant                'return value from find method
	'Public GotEMC As Variant           'True if either EMC or FAStT present
	'Public GotESS As Variant           'True if EMC/ESS or FAStT present
	Dim led As Integer                  'column number of last esoteric
	Dim MyCells As Range                'temp var
	Dim n As Integer                    'temp var
	Dim NewName As String               'new name for the EMC sheet
	Dim NewSheet As Worksheet
	'Public NumDISKs As Integer         'number of disk subsections
	Dim SectionName As String
	    
	SectionName = Sheet1.Name
	Call DelInt(Sheet1, numrows, 1)
	If SectionName = "DISKBSIZE" Then bSize = True
						'see if we have any esoterics on this sheet
	Set found = Sheet1.Rows(1).Find(Esoteric, LookAt:=xlPart)
	   
	If Not found Is Nothing Then
	   fed = found.Column
	   If Esoteric = "dac" Then NewName = "FASt" Else NewName = "EMC"
	   NewName = NewName & Right(SectionName, Len(SectionName) - 4)
						'sort the sheet to get the esoterics into one block
	   aa0 = ConvertRef(fed - 1)
	   Sheet1.Range(aa0 & "1:" & LastColumn & CStr(numrows)).Sort _
	      Key1:=Sheet1.Range(aa0 & "1"), Order1:=xlAscending, _
	      Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlLeftToRight
	   Set found = Sheet1.Rows(1).Find(Esoteric, LookAt:=xlPart)
	   fed = found.Column
						'create the new sheet
	   Sheet1.Copy after:=Sheets(Sheets.Count)
	   Sheets.Item(Sheets.Count).Name = NewName
	   Set NewSheet = Worksheets(NewName)
						'delete leading hdisks from new sheet
	   If fed > 2 Then
	      Set DelRange = NewSheet.Range("B1:" & ConvertRef(fed - 2) & "1")
	      DelRange.EntireColumn.Delete
	   End If
						'delete the esoterics from the DISK sheet
	   UserForm1.Label1.Caption = NewName
	   UserForm1.Repaint
	   Set DelRange = Sheet1.Range("B1:IU1")
	   For n = fed To 254
	      If Left(DelRange(1, n).Value, Len(Esoteric)) = Esoteric Then
		 led = led + 1
	      Else
		 Exit For
	      End If
	   Next n
	   led = led + fed
	   Set DelRange = Sheet1.Range(DelRange.Cells(1, fed - 1), DelRange.Cells(1, led - 1))
	   DelRange.EntireColumn.Delete
						'delete any trailing hdisks from the new sheet
	   If led < 253 Then
	      Set DelRange = NewSheet.Range(ConvertRef(led - fed + 2) & "1:IU1")
	      DelRange.EntireColumn.Delete
	   End If
	End If
						'do we still have data on the DISK sheet?
	If Sheet1.Range("B1") <> "" Then
	   Call GetLastColumn(Sheet1)
	   If Left(SectionName, 8) = "DISKXFER" And SVCTimes Then _
	      Call SVCgraph(numrows, SectionName, DiskSort And Not GotESS)
	   Call DiskGraphs(numrows, SectionName, DiskSort And Not GotESS)
					 'put in a totals column for DISK_SUMM
	   Sheet1.Range("IV1") = "Totals"
	   Set MyCells = Sheet1.Range("IV2:IV" & CStr(numrows))
	   MyCells(1, 1) = "=SUM(B2:" & LastColumn & "2)"
	   MyCells.FillDown
	   
	   If Left(SectionName, 8) = "DISKXFER" Then Call PP_DISKXFER(numrows, SectionName)
	End If
						'did we create an new sheet?
	If Not found Is Nothing Then
	   Call GetLastColumn(NewSheet)
	   NewSheet.Range("B1:IV1").HorizontalAlignment = xlRight
	   If InStr(NewName, "XFER") > 0 And SVCTimes Then _
	      Call SVCgraph(numrows, NewName, DiskSort)
	   Call DiskGraphs(numrows, NewName, DiskSort)
	End If
		     
End Sub
                                    'last mod v3.1.4
Sub PP_DISKXFER(numrows As Long, SectionName As String)
	Dim aa0 As String                       'temp var
	Dim Chart1 As ChartObject               'new chart object
	'Public Host As String                  'Hostname from AAA sheet
	'Public LastColumn As String            'last column letter
	Dim MyCells As Range                    'temp var
	Dim n As Integer                        'temp var
	'Public NumDisk as Integer              'number of disk subsections
	'Public RunDate As String               'NMON run date from AAA sheet
	Dim shDisk As Worksheet                 'pointer to DISK_SUMM sheet
	Dim shREAD As Worksheet                 'pointer to DISKREAD sheet
	Dim shWRITE As Worksheet                'pointer to DISKWRITE sheet
						    
	If Len(SectionName) = 8 Then            'If this is the first DISKXFER Sheet
						'Produce a sheet summarising data + I/O rates
	   UserForm1.Label1.Caption = "DISK_SUMM"
	   UserForm1.Repaint
	   Sheets.Add.Name = "DISK_SUMM"
	   Set shDisk = Worksheets("DISK_SUMM")
	   Set shREAD = Worksheets("DISKREAD")
	   Set shWRITE = Worksheets("DISKWRITE")
						'Copy timestamps + create colum heads etc.
	   shWRITE.Range("A1:A" & CStr(numrows)).Copy shDisk.Range("A1:A" & CStr(numrows))
	   aa0 = shWRITE.Range("A1").Value
	   Mid(aa0, 6) = "total"
	   shDisk.Range("A1") = aa0
	   shDisk.Range("B1") = "Disk Read kb/s"
	   shDisk.Range("C1") = "Disk Write kb/s"
	   shDisk.Range("D1") = "IO/sec"
	   Application.CutCopyMode = False
	   shDisk.Range("B2") = "=DISKREAD!IV2"
	   shDisk.Range("C2") = "=DISKWRITE!IV2"
	   shDisk.Range("D2") = "=DISKXFER!IV2"
	   shDisk.Range("B2:D" & CStr(numrows)).FillDown
	   shDisk.Columns("B:C").ColumnWidth = 12
						'Produce graph of data/IO rate on DISK_SUMM
	   If Graphs = "ALL" Or InStr(1, List, "DISK_SUMM") > 0 Then
	      Set MyCells = shDisk.Range("A1:C" & CStr(numrows))
	      Set Chart1 = shDisk.ChartObjects.Add(cLeft, cTop + rH * numrows, cWidth, cHeight)
	      Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, PlotBy:=xlColumns, _
	       CategoryLabels:=1, SeriesLabels:=1, Title:=shDisk.Range("A1") & " - " & RunDate
						   'apply style defaults
	   With Chart1.Chart
	      Call ApplyStyle(Chart1, 2, 3)
	      .SeriesCollection.NewSeries
	      With .SeriesCollection(3)
		 .AxisGroup = 2
		 .ChartType = xlLine
		 .Values = "=DISK_SUMM!R2C4:R" & CStr(numrows) & "C4"
		 .Name = "=DISK_SUMM!R1C4"
	      End With
	   End With
	   End If
					      
	Else                                'I already have a DISK_SUMM sheet so just update the data
	    NumDisk = NumDisk + 1
	    Set shDisk = Worksheets("DISK_SUMM")
	    n = Mid(SectionName, 9)
	    Set MyCells = shDisk.Range("B2:D" & CStr(numrows))
	    MyCells.Copy
	    MyCells.Offset(0, 3).PasteSpecial Paste:=xlPasteValues
	    aa0 = "READ" & n
	    MyCells(1, 1) = "=E2+DISK" & aa0 & "!IV2"
	    aa0 = "WRITE" & n
	    MyCells(1, 2) = "=F2+DISK" & aa0 & "!IV2"
	    aa0 = "XFER" & n
	    MyCells(1, 3) = "=G2+DISK" & aa0 & "!IV2"
	    MyCells.FillDown
	    MyCells.Copy
	    MyCells.PasteSpecial Paste:=xlPasteValues
	    shDisk.Range("E:G").Clear
	End If
	    
	shDisk.Move after:=Sheets(Sheets.Count)      'and move DISK_SUMM after me
		 
End Sub
                                       'last mod 2.9.0
Sub PP_ESS(numrows As Long, Sheet1 As Worksheet)
	'Public bSize As Variant                'True if DISKBSIZE present
	'Public LastColumn As String, ColNum as integer
	'Public NumDisk As Integer              'Number of disk subsections
	Dim aa0 As String                       'temp var
	Dim found As Range                      'result of find method
	Dim Host As String                      'name of host system from AAA sheet
	Dim MyCells As Range                    'temp var
	Dim nDisks As Integer                   'number of hdisks in a vpath
	Dim n As Integer, n1 As Integer         'loop counters
	Dim rRange As Range, bRange As Range
	Dim rString As String, bString As String
	Dim SectionName As String               'name of current sheet
	Dim sRange(19) As Range                  'list of search ranges
	Dim vname As String, hdisk As String, vTable As Range
	
	SectionName = Sheet1.Name
	Call DelInt(Sheet1, numrows, 1)
						'if this is the last sheet, do the analysis
	If SectionName = "ESSXFER" Then
	   UserForm1.Label1.Caption = "Analysing ESS data ... "
	   UserForm1.Repaint
						       'copy and set up sheets
	    Host = Worksheets("AAA").Range("Host").Value
	    Sheet1.Copy after:=Sheets(Sheets.Count)
	    Worksheets("ESSXFER (2)").Name = "ESSBUSY"
	    Worksheets("ESSBUSY").Range("A1") = "ESS %Busy " & Host
	    If bSize Then
	       Sheet1.Copy after:=Sheets(Sheets.Count)
	       Worksheets("ESSXFER (2)").Name = "ESSBSIZE"
	       Worksheets("ESSBSIZE").Range("A1") = "ESS xfer size (Kbytes) " & Host
	    End If
					       'define the search range
	    Set sRange(0) = Worksheets("DISKBUSY").Range("A1:IV1")
	    If NumDisk > 1 Then
	       For n = 1 To NumDisk - 1
		Set sRange(n) = Worksheets("DISKBUSY" & CStr(n)).Range("A1:IV1")
	       Next n
	    End If
		
	    Set vTable = Worksheets("BBBE").Range("vPaths")
	    Set MyCells = Worksheets("ESSBUSY").Range("B1:IV1")
	    If bSize Then
	       Set rRange = Worksheets("ESSBSIZE").Range("B2:IV" & CStr(numrows))
	    End If
	    Set bRange = Worksheets("ESSBUSY").Range("B2:IV" & CStr(numrows))
						'and now build the formulas
	    For ColNum = 1 To 255               'for each vpath
	       rString = "=("
	       bString = "=("
	       If MyCells(1, ColNum) = "" Then Exit For
	       vname = MyCells(1, ColNum)
	       nDisks = Application.WorksheetFunction.VLookup(vname, vTable, 2, False)
						'for each hdisk in the vpath
	       If nDisks > 0 Then
		  For n = 1 To nDisks
		     hdisk = Application.WorksheetFunction.VLookup(vname, vTable, 2 + n, False)
						'find out where the data is
		     For n1 = 0 To NumDisk - 1
			 Set found = sRange(n1).Find(hdisk, LookAt:=xlWhole)
			 If Not found Is Nothing Then Exit For
		     Next n1
	      
		     If found Is Nothing Then
			MsgBox ("Error - unable to locate data for " & hdisk)
		     Else
						 'add to the formula strings
			aa0 = found.AddressLocal(False, True)
			aa0 = "!" & Left(aa0, Len(aa0) - 1) & "2"
			If n1 > 0 Then aa0 = CStr(n1) & aa0
			rString = rString & "+DISKBSIZE" + aa0
			bString = bString & "+DISKBUSY" + aa0
		    End If
		  Next n
							'last hdisk - update cell contents
		  If bSize Then
		     rString = rString & ")/" & CStr(nDisks)
		     rRange(1, ColNum) = rString
		  End If
		  bString = bString & ")/" & CStr(nDisks)
		  bRange(1, ColNum) = bString
		  Sheet1.Range("B1").Cells(numrows + 6, ColNum) = nDisks 'for SVCgraph
	       End If
	    Next ColNum
						 'last vpath - complete sheets
	    If bSize Then
	       UserForm1.Label1.Caption = "ESSBSIZE"
	       UserForm1.Repaint
	       rRange.FillDown
	       rRange.Copy
	       rRange.PasteSpecial Paste:=xlPasteValues
	       Application.CutCopyMode = False
	       rRange.NumberFormat = "0.0"
	       Call DiskGraphs(numrows, "ESSBSIZE", DiskSort)
	    End If
	    UserForm1.Label1.Caption = "ESSBUSY"
	    UserForm1.Repaint
	    bRange.FillDown
	    bRange.Copy
	    bRange.PasteSpecial Paste:=xlPasteValues
	    Application.CutCopyMode = False
	    bRange.NumberFormat = "0.0"
	    Call DiskGraphs(numrows, "ESSBUSY", DiskSort)
	       
	   If SVCTimes Then Call SVCgraph(numrows, SectionName, DiskSort)
	   Call DiskGraphs(numrows, "ESSXFER", DiskSort)
	Else
	   Call DiskGraphs(numrows, SectionName, DiskSort)
	End If
	    
End Sub
                                       'last mod v3.1.0
Sub PP_FILE(numrows As Long, Sheet1 As Worksheet)
	Dim MyCells As Range                    'temp var
	Dim Chart1 As ChartObject               'new chart object
	'Public RunDate As String               'NMON run date from AAA sheet
	Dim SectionName As String               'name of current sheet
	    
	SectionName = Sheet1.Name
	Call DelInt(Sheet1, numrows, 1)
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
	
						'Add Avg/Wavg/Max lines
	Set MyCells = avgmax(numrows, Sheet1, 0)
						'Produce a graph of readch/writech rates
	Set MyCells = Sheet1.Range("E2:F" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	   Title:="Kernel Read/Write System Calls " & Host & "  " & RunDate
	With Chart1.Chart                    'apply customisation
	    .SeriesCollection(1).Name = "readch/sec"
	    .SeriesCollection(2).Name = "writech/sec"
	    .SeriesCollection(1).XValues = "=FILE" & "!R3C1:R" & CStr(numrows) & "C1"
	    Call ApplyStyle(Chart1, 1, 2)
	End With
						'Produce a graph of i-node data
	Set MyCells = Sheet1.Range("A1:D" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	   CategoryLabels:=1, SeriesLabels:=1, HasLegend:=True, _
	   Title:="Kernel Filesystem Functions " & Host & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   Call ApplyStyle(Chart1, 1, 3)
	End With
	    
End Sub
                                       'last mod v3.1.0
Sub PP_FRCA(numrows As Long, Sheet1 As Worksheet)     'last mod v3.1.0
	Dim Chart1 As ChartObject               'new chart object
	Dim MyCells As Range                    'range for charting
	'Public RunDate As String               'NMON run date from AAA sheet
	
	Call DelInt(Sheet1, numrows, 1)
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
						'Produce a graph of cache hits stats
	Set MyCells = Sheet1.Range("F1:F" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, cTop + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	   SeriesLabels:=1, HasLegend:=False, CategoryTitle:="Time of Day", _
	   ValueTitle:=Sheet1.Range("F1").Value, _
	   Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=FRCA" & "!R2C1:R" & CStr(numrows) & "C1"
	   .Axes(xlValue).MaximumScale = 1
	   Call ApplyStyle(Chart1, 1, 1)
	   .Axes(xlValue).TickLabels.NumberFormat = "0%"
	End With
	  
End Sub
                                       'last mod v3.1.0
Sub PP_IOADAPT(numrows As Long, Sheet1 As Worksheet)
	Dim aa0 As String                       'temp var
	Dim Chart1 As ChartObject               'new chart object
	Dim CurCol As Integer                   'loop counter
	Dim Graphdata As Range                  'range for tps graph
	'Public LastColumn as String, ColNum as Integer
	Dim MyCells As Range                    'temp var
	Dim NumAdapt As Integer                 'number of I/O adapters on sheet
	'Public RunDate As String               'NMON run date from AAA sheet
	Dim SectionName As String               'name of current sheet
	Dim tpsCol As Range                     'pointer for move
	    
	SectionName = Sheet1.Name
	Call DelInt(Sheet1, numrows, 1)
						'Put in the formulas for avg/max
	Set Graphdata = avgmax(numrows, Sheet1, 0)
	
	NumAdapt = ColNum / 3
	Set MyCells = Sheet1.Range("A1")
						'move columns around for graphing
	MyCells(1, 1) = MyCells(1, 1).Value & " (KB/s)" 'title
	For CurCol = 2 To (NumAdapt * 2) + 1 Step 2
	    aa0 = MyCells(1, CurCol).Value
	    MyCells(1, CurCol) = Left(aa0, Len(aa0) - 5) 'strip off KB/s from read
	    aa0 = MyCells(1, CurCol + 1).Value
	    MyCells(1, CurCol + 1) = Left(aa0, Len(aa0) - 5)      'strip off KB/s from write
	    aa0 = Sheet1.Range("A1").Item(1, CurCol + 2).Address(True, False, xlA1)
	    aa0 = Left(aa0, InStr(1, aa0, "$") - 1)
	    Set tpsCol = Sheet1.Range(aa0 & "1.." & aa0 & CStr(numrows))
	    tpsCol.Copy MyCells(1, ColNum + 1)   'move tps column to the end
	    tpsCol.EntireColumn.Delete
	Next CurCol
	Sheet1.Range("B:" & LastColumn).ColumnWidth = 12
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
						'produce avg/max graph
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=Graphdata, PlotBy:=xlRows, _
	SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	    .SeriesCollection(1).XValues = "=" & SectionName & "!R1C2:R1C" & CStr(ColNum)
	    Call ApplyStyle(Chart1, 0, 3)
	End With
						'produce data rate graph
	aa0 = ConvertRef(NumAdapt * 2)
	Set Graphdata = Sheet1.Range("A1:" & aa0 & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=Graphdata, Gallery:=xlArea, _
	    CategoryLabels:=1, SeriesLabels:=1, _
	    Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	    Call ApplyStyle(Chart1, 2, 99)
	End With
						'produce tps graph
	aa0 = Sheet1.Range("A1").Item(1, ColNum - (NumAdapt - 1)).Address(True, False, xlA1)
	aa0 = Left(aa0, InStr(1, aa0, "$") - 1)
	Set Graphdata = Sheet1.Range(LastColumn & "1:" & aa0 & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop3 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=Graphdata, Gallery:=xlArea, _
	    Title:="Disk Adapter " & Host & " (tps) - " & RunDate, _
	    SeriesLabels:=1, HasLegend:=True
	
	With Chart1.Chart                    'apply customisation
	    .SeriesCollection(1).XValues = "=IOADAPT" & "!R2C1:R" & CStr(numrows) & "C1"
	    Call ApplyStyle(Chart1, 2, 99)
	End With
	    
End Sub
                                       'last mode v3.1.5
Sub PP_JFS(numrows As Long, Sheet1 As Worksheet)
	Dim aa0 As String, aa1 As String
	Dim Chart1 As ChartObject              'new chart object
	Dim ColNum As Integer
	Dim MyCells As Range                   'range for charting
	'Public RunDate As String               'NMON run date from AAA sheet
	Dim Source As Range
	Dim Target As Range
	
	Call DelInt(Sheet1, numrows, 1)
						'do we have a /proc heading
	Set Target = Sheet1.Rows(1).Find("/proc", LookAt:=xlWhole)
						'if so, delete the heading (but not the data)
	If Not Target Is Nothing Then
	   aa0 = Target.AddressLocal(False, False)
	   aa1 = Left(aa0, Len(aa0) - 1)
	   ColNum = ConvertRef(aa1)
	   aa1 = ConvertRef(ColNum)
	   Set Source = Sheet1.Range(aa1 & "1:IV1")
	   Source.Copy Sheet1.Range(aa0 & ":IU1")
	   Call GetLastColumn(Sheet1)
	End If
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
						'produce avg/max graph
	Set MyCells = avgmax(numrows, Sheet1, 0)
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, PlotBy:=xlRows, SeriesLabels:=1, _
	    Title:=Sheet1.Range("A1").Value & "  " & RunDate
	With Chart1.Chart                      'apply customisation
	    .SeriesCollection(1).XValues = Sheet1.Range("B1:" & LastColumn & "1")
	    .Axes(xlValue).MaximumScale = 100
	    Call ApplyStyle(Chart1, 0, 3)
	End With
End Sub
                                       'new for v3.0.7
Sub PP_LPAGE(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject               'new chart object
	Dim MyCells As Range                    'range used for charting
	'Public RunDate As String               'NMON run date from AAA sheet
	
	Call DelInt(Sheet1, numrows, 1)
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
					       'graph showing %breakdown
	Set MyCells = Sheet1.Range("B1..C" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=LARGEPAGE" & "!R2C1:R" & CStr(numrows) & "C1"
	'   .Axes(xlValue).MaximumScale = 100
	   Call ApplyStyle(Chart1, 2, 2)
	End With
End Sub
                                       'v3.2.4
Sub PP_LPAR(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject               'new chart object
	Dim MyCells As Range                    'range used for charting
	'Public RunDate As String               'NMON run date from AAA sheet
	
	Call DelInt(Sheet1, numrows, 1)
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
					       'produce a graph of physcpu vs entitled
	If Linux Then
	    Set MyCells = Sheet1.Range("B1:B" & CStr(numrows) & ",I1:I" & CStr(numrows))
	Else
	    Set MyCells = Sheet1.Range("B1:B" & CStr(numrows) & ",F1:F" & CStr(numrows))
	End If
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	   SeriesLabels:=1, Title:="Physical CPU vs Entitlement - " & Host & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	     .SeriesCollection(1).XValues = "=LPAR" & "!R2C1:R" & CStr(numrows) & "C1"
	     Call ApplyStyle(Chart1, 1, 2)
	End With
	If Linux Then Exit Sub
					       'graph showing use in pool
	Sheet1.Range("L1").Value = "OtherLPARs"
	Sheet1.Range("L2").Value = "=E2-B2-H2"
	Set MyCells = Sheet1.Range("L2:L" & CStr(numrows))
	MyCells.FillDown
	MyCells.NumberFormat = "0.00"
	
	Set MyCells = Sheet1.Range("B1:B" & CStr(numrows) & ",L1:L" & CStr(numrows) & ",H1:H" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	SeriesLabels:=1, Title:="Shared Pool Utilisation - " & Host & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=LPAR" & "!R2C1:R" & CStr(numrows) & "C1"
	   Call ApplyStyle(Chart1, 2, 3)
	End With
					       'Wavg data
	Set MyCells = avgmax(numrows, Sheet1, 0)

End Sub
                                       'last mod v3.0.3
Sub PP_MEM(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject               'new chart object
	Dim MyCells As Range                    'range used for charting
	'Public RunDate As String               'NMON run date from AAA sheet
	    
	Sheet1.Range("B1:O1").Columns.AutoFit
	Call DelInt(Sheet1, numrows, 1)
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
						'Produce a graph of free memory
	If Linux Then
	    Set MyCells = Sheet1.Range("F1:F" & CStr(numrows))
	Else
	    Set MyCells = Sheet1.Range("D1:D" & CStr(numrows))
	End If
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, cTop + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	   SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & "  " & RunDate
	   
	With Chart1.Chart                       'apply customisation
	     .SeriesCollection(1).XValues = "=MEM" & "!R2C1:R" & CStr(numrows) & "C1"
	     Call ApplyStyle(Chart1, 1, 1)
	End With
						'produce a graph of real memory                                                                       'Produce a graph of free memory
	If Linux Then
	    Set MyCells = Sheet1.Range("B1:B" & CStr(numrows))
	Else
	    Set MyCells = Sheet1.Range("F1:F" & CStr(numrows))
	End If
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, dlHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	   SeriesLabels:=1, HasLegend:=True
						'apply customisation
	With Chart1.Chart
	     .SeriesCollection(1).XValues = "=MEM" & "!R2C1:R" & CStr(numrows) & "C1"
	     Call ApplyStyle(Chart1, 2, 1)
	     .HasTitle = False
	End With
End Sub
                                       'new for v3.1.0
Sub PP_MEMNEW(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject               'new chart object
	Dim MyCells As Range                    'range used for charting
	'Public RunDate As String               'NMON run date from AAA sheet
	Dim MyTitle As String
	
	Call DelInt(Sheet1, numrows, 1)
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
	
	MyTitle = Sheet1.Range("A1").Value
	Mid(MyTitle, 8, 3) = "Use"
						'graph showing %breakdown
	Set MyCells = Sheet1.Range("B1..D" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	SeriesLabels:=1, Title:=MyTitle & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=MEMNEW" & "!R2C1:R" & CStr(numrows) & "C1"
	   .Axes(xlValue).MaximumScale = 100
	   Call ApplyStyle(Chart1, 2, 2)
	End With
End Sub
                                       'last mod v3.1.0
Sub PP_MEMUSE(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject              'new chart object
	Dim MyCells As Range                   'range used for charting
	'Public RunDate As String              'NMON run date from AAA sheet
	
	Call DelInt(Sheet1, numrows, 1)
					       'set up column headings + formulas for graphing
	Sheet1.Range("G1").Value = "%comp"
	Sheet1.Range("G2").Value = "=100-MEM!B2-B2"
	Set MyCells = Sheet1.Range("G2:G" & CStr(numrows))
	MyCells.FillDown
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
					       'Produce a graph of %numperm
	Set MyCells = Sheet1.Range("B1:D" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
		SeriesLabels:=1, Title:="VMTUNE Parameters " & Host & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=MEMUSE" & "!R2C1:R" & CStr(numrows) & "C1"
	   .Axes(xlValue).MaximumScale = 100
	   Call ApplyStyle(Chart1, 1, 3)
	   .Axes(xlValue).HasMajorGridlines = False
	   .SeriesCollection(2).Border.LineStyle = xlDot
	   .SeriesCollection(3).Border.LineStyle = xlDot
	   .SeriesCollection(3).Border.ColorIndex = 26   'Magenta
	   .SeriesCollection.NewSeries
	   With .SeriesCollection(4)
	     .Values = "=MEMUSE!R2C7:R" & CStr(numrows) & "C7"
	     .Name = "%comp"
	   End With
	End With
	If SheetExists("MEMNEW") Then Exit Sub
						'graph showing %comp and %file
	Set MyCells = Sheet1.Range("G1..G" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=MEMUSE" & "!R2C1:R" & CStr(numrows) & "C1"
	   .SeriesCollection.NewSeries
	   With .SeriesCollection(2)
	     .Values = "=MEMUSE!R2C2:R" & CStr(numrows) & "C2"
	     .Name = "%file"
	   End With
	   .Axes(xlValue).MaximumScale = 100
	   Call ApplyStyle(Chart1, 2, 2)
	End With
End Sub
                                       'last mod v3.2.5
Sub PP_NET(numrows As Long, Sheet1 As Worksheet)
	Dim aa0 As String                       'temp var
	Dim Chart1 As ChartObject               'new chart object
	'Public ColNum As Integer               'last column number
	Dim CurCol As Integer                   'loop counter
	Dim Graphdata As Range                  'saved range from avgmax function
	Dim incr As Integer                     'number of network adapters
	'Public LastColumn As String            'last column letter
	Dim MyCells As Range                    'temp var
	Dim n As Integer                        'number of columns to traverse
	'Public RunDate As String               'NMON run date from AAA sheet
	
	Call DelInt(Sheet1, numrows, 1)
						'Put in the formulas for avg/max
	Set Graphdata = avgmax(numrows, Sheet1, 0)
	n = ColNum + 1
	incr = (n - 2) / 2
	Set MyCells = Sheet1.Range("A1")
						'set up column headings + formulas for graphing
	MyCells(1, 1) = MyCells(1, 1).Value & " (KB/s)"
	For CurCol = n To (n + incr - 1)
	   aa0 = MyCells(1, CurCol - incr * 2).Value
	   MyCells(1, CurCol - incr * 2) = Left(aa0, Len(aa0) - 5)
	   aa0 = MyCells(1, CurCol - incr).Value
	   aa0 = Left(aa0, Len(aa0) - 5)       'strip off KB/s from write
	   MyCells(1, CurCol - incr) = aa0
	   aa0 = Left(aa0, Len(aa0) - 5) & "total"
	   MyCells.Item(1, CurCol) = aa0       'add total
	   MyCells.Item(2, CurCol) = "=" & ConvertRef(CurCol - incr - 1) & "2+" & ConvertRef(CurCol - incr * 2 - 1) & "2"
	Next CurCol
	MyCells(1, n + incr) = "Total-Read"
	MyCells(2, n + incr) = "=SUM(B2:" & ConvertRef(incr) & "2)"
	MyCells(1, n + incr + 1) = "Total-Write (-ve)"
	MyCells(2, n + incr + 1) = "=-SUM(" & ConvertRef(incr + 1) & "2.." & ConvertRef(n - 2) & "2)"
	Sheet1.Range(Sheet1.Cells(2, n), Sheet1.Cells(numrows, n + incr + 1)).FillDown
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
						'Produce total reads/writes graph
	Set MyCells = Sheet1.Range(Sheet1.Cells(1, n + incr), Cells(numrows, n + incr + 1))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	   PlotBy:=xlColumns, SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & " - " & RunDate
	     
	With Chart1.Chart                       'apply customisation
	   .SeriesCollection(1).XValues = "=NET!R2C1:R" & CStr(numrows) & "C1"
	   Call ApplyStyle(Chart1, 2, 2)
	   .ChartType = xlArea
	   .Axes(xlValue).MinimumScaleIsAuto = True
	   .Axes(xlCategory).TickLabelPosition = xlLow
	End With
						'produce avg/max graph
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=Graphdata, PlotBy:=xlRows, _
	   SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=NET!R1C2:R1C" & CStr(ColNum)
	   Call ApplyStyle(Chart1, 0, 3)
	End With
						'produce adapter by ToD graph
	Set MyCells = Sheet1.Range("A1:" & LastColumn & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop3 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	   PlotBy:=xlColumns, CategoryLabels:=1, SeriesLabels:=1, _
	   Title:=Sheet1.Range("A1").Value & "  " & RunDate
	   
	With Chart1.Chart                       'apply customisation
	  .SeriesCollection(1).XValues = "=NET!R2C1:R" & CStr(numrows) & "C1"
	  Call ApplyStyle(Chart1, 2, 99)
	End With

End Sub
                                       'last mod v3.1.0
Sub PP_NETP(numrows As Long, Sheet1 As Worksheet)
	Dim aa0 As String                       'temp var
	Dim Chart1 As ChartObject               'new chart object
	'Public ColNum As Integer               'last column number
	Dim CurCol As Integer                   'loop counter
	Dim Graphdata As Range                  'range for graphing
	'Public Host As String                  'Hostname from AAA sheet
	'Public LastColumn As String            'last column letter
	Dim MyCells As Range                    'temp var
	Dim n As Integer                        'number of columns at the start
	'Public RunDate As String               'NMON run date from AAA sheet
	
	Call DelInt(Sheet1, numrows, 1)
	n = ColNum
						'set up col headings + formulas for packet size
	Set MyCells = Sheet1.Range("A1")
	For CurCol = ColNum + 1 To 2 * ColNum - 1
	   aa0 = MyCells(1, CurCol - ColNum + 1).Value
	   MyCells.Item(1, CurCol) = Left(aa0, InStr(1, aa0, "-") + 1) & "size"
	Next CurCol
	MyCells.Item(2, ColNum + 1) = "=IF(B2>0,NET!B2/B2*1024,0)"
	aa0 = ConvertRef(ColNum)                'save Column pointer
	Call GetLastColumn(Sheet1)
	Set MyCells = Sheet1.Range(aa0 & "2:" & LastColumn & CStr(numrows))
	If LastColumn <> "B" Then MyCells.FillRight
	MyCells.FillDown
	MyCells.NumberFormat = "0.0"
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
						'produce avg/max graph
	Set MyCells = avgmax(numrows, Sheet1, 0)
	aa0 = ConvertRef(n)
	Set Graphdata = Sheet1.Range(aa0 & CStr(numrows + 2) & ":" & LastColumn & CStr(numrows + 4))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=Graphdata, PlotBy:=xlRows, _
	   Title:="Network Packet Size (bytes) " & Host & "  " & RunDate
	With Chart1.Chart                       'apply customisation
	  .SeriesCollection(1).Name = "Avg."
	  .SeriesCollection(2).Name = "WAvg."
	  .SeriesCollection(3).Name = "Max."
	  .SeriesCollection(1).XValues = "=NETPACKET!R1C" & CStr(n + 1) & ":R1C" & CStr(n * 2)
	  Call ApplyStyle(Chart1, 0, 3)
	End With
						'Produce graph of packets/s
	Set Graphdata = Sheet1.Range("A1:" & ConvertRef(n - 1) & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=Graphdata, Gallery:=xlArea, _
	   SeriesLabels:=1, CategoryLabels:=1, _
	   Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	  Call ApplyStyle(Chart1, 2, 99)
	End With
End Sub
                                    'new for v3.1.0
Sub PP_NFS(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject               'new chart object
	Dim MyCells As Range                    'temp var
	Dim SectionName As String               'sheet name
	
	Call DelInt(Sheet1, numrows, 1)
	If Application.WorksheetFunction.Max(Sheet1.Range("B2:S" & CStr(numrows))) = 0 Then
	   Application.DisplayAlerts = False
	   Sheet1.Delete
	   Application.DisplayAlerts = True
	   Exit Sub
	End If
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
					       'graph reads/writes
	SectionName = Sheet1.Name
	Set MyCells = Sheet1.Range("H1:I" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	   SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & " " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=" & SectionName & "!R2C1:R" & CStr(numrows) & "C1"
	   Call ApplyStyle(Chart1, 2, 2)
	End With
End Sub
                                    'last mod v3.1.4
Sub PP_PAGE(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject               'new chart object
	Dim MyCells As Range                    'range used for charting
	'Public RunDate As String               'NMON run date from AAA sheet
	
	Call DelInt(Sheet1, numrows, 1)
	Sheet1.Range("J1").Value = "fsin"
	Sheet1.Range("K1").Value = "fsout"
	Sheet1.Range("J2").Value = "=C2-E2"
	Sheet1.Range("K2").Value = "=D2-F2"
	Sheet1.Range("L1").Value = "sr/fr"
	Sheet1.Range("L2").Value = "=IF(G2>0,H2/G2,0)"
	Set MyCells = Sheet1.Range("J2:L" & CStr(numrows))
	MyCells.FillDown
	MyCells.NumberFormat = "0.0"
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
						'graph the paging rates
	Set MyCells = Sheet1.Range("E1:F" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	   SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & " (pgspace)  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=PAGE" & "!R2C1:R" & CStr(numrows) & "C1"
	   Call ApplyStyle(Chart1, 2, 2)
	End With
						'graph the filespace rates
	Set MyCells = Sheet1.Range("J1:K" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlArea, _
	   SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & " (filesystem)  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=PAGE" & "!R2C1:R" & CStr(numrows) & "C1"
	   Call ApplyStyle(Chart1, 2, 2)
	End With
						'graph the sr/fr rate
	Set MyCells = Sheet1.Range("L1:L" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop3 + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	   SeriesLabels:=1, Title:="Page scan:free ratio " & Host & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	     .SeriesCollection(1).XValues = "=PAGE" & "!R2C1:R" & CStr(numrows) & "C1"
	     Call ApplyStyle(Chart1, 1, 1)
	      .HasLegend = False
	End With
	      
End Sub
                                       'new for v3.1.5
Sub PP_PAGING(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject
	Dim MyCells As Range
	
	Call DelInt(Sheet1, numrows, 1)
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
						'produce line graph
	Set MyCells = Sheet1.Range("A1:" & LastColumn & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, cTop + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	    Title:=Sheet1.Range("A1").Value & "  " & RunDate, CategoryLabels:=1, SeriesLabels:=1, HasLegend:=True
						'apply customisation
	With Chart1.Chart
	    Call ApplyStyle(Chart1, 1, 99)
	End With
End Sub
                                       'last mod v3.1.0
Sub PP_PROC(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject
	Dim MyCells As Range
	
	Call DelInt(Sheet1, numrows, 1)
	Sheet1.Range("B1") = "RunQueue"
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
	
						'produce RunQueue/swap-ins graph
	Set MyCells = Sheet1.Range("B1:C" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	    SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).XValues = "=PROC" & "!R2C1:R" & CStr(numrows) & "C1"
	   Call ApplyStyle(Chart1, 1, 2)
	End With
						'Produce a graph of pswitch/syscalls rates
	Set MyCells = Sheet1.Range("D1:E" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	      SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).Name = "pswitch/sec"
	   .SeriesCollection(2).Name = "syscalls/sec"
	   .SeriesCollection(1).XValues = "=PROC" & "!R2C1:R" & CStr(numrows) & "C1"
	   Call ApplyStyle(Chart1, 1, 2)
	End With
						'Produce a graph of forks/execs
	Set MyCells = Sheet1.Range("H1:I" & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop3 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	      SeriesLabels:=1, Title:=Sheet1.Range("A1").Value & "  " & RunDate
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection(1).Name = "forks/sec"
	   .SeriesCollection(2).Name = "execs/sec"
	   .SeriesCollection(1).XValues = "=PROC" & "!R2C1:R" & CStr(numrows) & "C1"
	   Call ApplyStyle(Chart1, 1, 2)
	End With
End Sub
                                       'last mod v3.1.0
Sub PP_PROCAIO(numrows As Long, Sheet1 As Worksheet)
	Dim Chart1 As ChartObject               'new chart object
	'Public ColNum as Integer
	'Public Host As String                  'Hostname from AAA sheet
	'Public LastColumn As String
	Dim MyCells As Range                    'range for charting
	'Public RunDate As String               'NMON run date from AAA sheet
	Dim SectionName As String
	    
	SectionName = Sheet1.Name
	Call DelInt(Sheet1, numrows, 1)
	If Not SheetExists("CPU_ALL") Then Exit Sub
						'add syscpu column
	Sheet1.Range("E1").Value = "syscpu"
	Sheet1.Range("E2").Value = "=D2/CPU_ALL!G2"
	Set MyCells = Sheet1.Range("E2:E" & CStr(numrows))
	Sheet1.Range("E2:E" & CStr(numrows)).NumberFormat = "#,##0.0"
	MyCells.FillDown
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
						'produce avg/max graph
	Set MyCells = avgmax(numrows, Sheet1, 0)
	Set MyCells = Sheet1.Range("A" & CStr(numrows + 2) & ":C" & CStr(numrows + 4))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, PlotBy:=xlRows, SeriesLabels:=1, _
	    Title:=Sheet1.Range("A1").Value & "  " & RunDate
	With Chart1.Chart                       'apply customisation
	    .SeriesCollection(1).XValues = "=" & SectionName & "!R1C2:R1C" & CStr(ColNum)
	    Call ApplyStyle(Chart1, 0, 3)
	End With
						'produce line graph
	Set MyCells = Union(Sheet1.Range("A1:A" & CStr(numrows)), _
	    Sheet1.Range("C1:C" & CStr(numrows)), Sheet1.Range("E1:E" & CStr(numrows)))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	    CategoryLabels:=1, SeriesLabels:=1, PlotBy:=xlColumns, _
	    Title:=Sheet1.Range("A1").Value & "  " & RunDate
	With Chart1.Chart
	   With .SeriesCollection(2)
		.AxisGroup = 2
		.MarkerStyle = xlNone
	   End With
	   Call ApplyStyle(Chart1, 1, 2)
	   .Axes(xlValue).HasMajorGridlines = False
	   .Axes(xlValue, xlPrimary).HasTitle = True
	   .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "#active processes"
	   .Axes(xlValue, xlSecondary).HasTitle = True
	   .Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "%cpu used by aio"
	End With
End Sub
                                       'v3.2.6
Sub PP_SUMMARY(numrows As Long, Sheet1 As Worksheet)
' place holder
End Sub
                                       'last mod v3.2.6
Sub PP_TOP(numrows As Long, Sheet1 As Worksheet)
	'Public NumCPUs As Integer              'Number of CPU sections
	Dim aa0 As String                       'temp var
	Dim aa1 As String                       'temp var
	Dim aa2 As String                       'string representing NumCPUs
	Dim Chart1 As ChartObject               'new chart object
	Dim Cl As String                        'Column letter for IntervalCPU%
	Dim Cw As String                        'Column letter for WSet
	Dim Cmd As String                       'current command
	Dim Cn As Integer                       'column number for IntervalCPU%
	Dim CurRow As Long                      'loop counter
	Dim MyCells As Range                    'temp var
	Dim MyRow As Long                       'temp var
	Dim NumCmds As Integer                  'number of unique commands
	Dim oldCmd As String                    'previous command
	Dim AvgCPUs As Integer
	Dim sRow As Long                        'start row of command block
	Dim sTRow As Long                       'start row of interval block
	Dim TVal As String                      'current time value
	Dim oldTVal As String                   'previous time value
	
	If numrows < 3 Then
	   Application.DisplayAlerts = False
	   Sheet1.Delete
	   Application.DisplayAlerts = True
	   Exit Sub
	End If
	UserForm1.Label1.Caption = "TOP - reorganising data..."
	UserForm1.Repaint
						'check that the headers have been correctly sorted
	If Sheet1.Range("A1").Value = "%CPU Utilisation" Then
	   Sheet1.Rows(1).EntireRow.Delete
	Else                                    'it must be at the bottom
	   Sheet1.Range("A" & CStr(numrows - 1)).EntireRow.Delete
	   Sheet1.Range("A1").EntireRow.Insert
	   Sheet1.Range("A" & CStr(numrows) & ":IV" & CStr(numrows)).Cut
	   Sheet1.Range("A1").Paste
	End If
	numrows = numrows - 1
						'Start out by putting the data in a more reasonable order
	Sheet1.Columns("C").Insert Shift:=xlToRight
	Set MyCells = Sheet1.Range("A2:A" & CStr(numrows))
	MyCells.Copy Sheet1.Range("C2:C" & CStr(numrows))
	Sheet1.Columns("A").EntireColumn.Delete
	
	If topas Then
	   Sheet1.Columns("G:H").Insert Shift:=xlToRight
	   Set MyCells = Sheet1.Range("C1:C" & CStr(numrows))
	   MyCells.Copy Sheet1.Range("G1:G" & CStr(numrows))
	   Sheet1.Columns("C").EntireColumn.Delete
	End If
						'quick fix for Excel formatting problem
	Sheet1.Range("B1").Value = "PID"
						'delete unwanted intervals
	If first > 1 Or last < 9999 Then
	   Sheet1.Columns("A:Z").Sort Key1:=Sheet1.Range("A2"), Order1:=xlAscending, _
	      Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
	      TVal = Sheet1.Range("A" & CStr(numrows)).Value
	      MyRow = CInt(Right(TVal, 4))
	   If last < MyRow Then
	      TVal = "T" & Format(last + 1, "0000")
	      Set MyCells = Sheet1.UsedRange.Find(TVal, LookAt:=xlWhole)
	      If Not (MyCells Is Nothing) Then
		 MyRow = MyCells.Row
		 Sheet1.Range("A" & CStr(MyRow) & ":A" & CStr(numrows)).EntireRow.Delete
		 numrows = MyRow - 1
	      End If
	   End If
	   If first > 1 Then
	      TVal = "T" & Format(first + 1, "0000")
	      Set MyCells = Sheet1.UsedRange.Find(TVal, LookAt:=xlWhole)
	      If Not (MyCells Is Nothing) Then
	      If MyCells.Row > 2 Then
		 MyRow = MyCells.Row - 1
		 Sheet1.Range("A2:A" & CStr(MyRow)).EntireRow.Delete
		 numrows = numrows - MyRow + 1
	      End If
	      End If
	   End If
	End If
	TopRows = numrows
						'sort the data into Command name order
	Sheet1.Columns("A:Z").Sort Key1:=Sheet1.Range("M2"), Order1:=xlAscending, _
	   Key2:=Sheet1.Range("A2"), Order2:=xlAscending, Header:=xlYes, _
	   OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
	Sheet1.Columns(13).AutoFit
						'find a free column
	Set MyCells = Sheet1.Range("A1")
	For Cn = 14 To 254
	    If MyCells(1, Cn) = "" Then Exit For
	Next Cn
	Cl = MyCells.Item(1, Cn).Address(True, False, xlA1)
	Cl = Left(Cl, InStr(1, Cl, "$") - 1)
	Cw = MyCells.Item(1, Cn + 1).Address(True, False, xlA1)
	Cw = Left(Cw, InStr(1, Cw, "$") - 1)
	
	MyCells(1, Cn) = "IntervalCPU%"
	MyCells(1, Cn + 1) = "WSet"
						'and now produce CPU totals for each command
	If dLPAR Then
	    aa1 = ",CPU_ALL!A$2:G$" & CStr(CPUrows) & ",7)"
	    aa2 = "VLOOKUP(A2" & aa1
	Else
	    aa2 = Worksheets("CPU_ALL").Range("G2")
	End If
	AvgCPUs = Application.WorksheetFunction.Average(Worksheets("CPU_ALL").Range("G2:G" & CStr(CPUrows)))
						'calc totals for all single-process cmds
	MyCells(2, Cn) = "=IF(A2=A3,IF(M2=M3,"" "",C2/" & aa2 & "),C2/" & aa2 & ")"
	MyCells(2, Cn + 1) = "=IF(A2=A3,IF(M2=M3,"" "",H2+I2),H2+I2)"
	If numrows > 2 Then
	   MyCells.Range(Cells(2, Cn), Cells(numrows, Cn + 1)).FillDown
	   MyCells.Range(Cells(2, Cn), Cells(numrows, Cn + 1)).Copy
	   MyCells.Range(Cells(2, Cn), Cells(numrows, Cn + 1)).PasteSpecial Paste:=xlPasteValues
	End If
					
	sRow = 2
	sTRow = 2
	NumCmds = 1
	oldCmd = MyCells(sRow, 13).Value        'Commands are in column "M" = 13
	oldTVal = MyCells(sRow, 1).Value
	For CurRow = sRow To 65535
						'generate sub-totals by time interval
	   TVal = MyCells(CurRow, 1).Value
	   If TVal <> oldTVal Then
	      If (sTRow + 1) <> CurRow Then
		 If dLPAR Then aa2 = "VLOOKUP(A" & CStr(CurRow - 1) & aa1
		 MyCells(CurRow - 1, Cn) = "=SUM(C" & CStr(sTRow) & ":C" & CStr(CurRow - 1) & ")/" & aa2
		 MyCells(CurRow - 1, Cn + 1) = "=SUM(I" & CStr(sTRow) & ":I" & CStr(CurRow - 1) & ")+H" & CStr(sTRow)
	      End If
	      sTRow = CurRow
	      oldTVal = TVal
	   End If
	       
	   Cmd = MyCells.Item(CurRow, 13).Value
	   If Cmd <> oldCmd Then
						'create table entry for this command
	      MyRow = numrows + NumCmds + 2
	      NumCmds = NumCmds + 1
	      MyCells(MyRow, 2) = oldCmd
	      MyCells(MyRow, 3) = _
		 "=SUM(C" & CStr(sRow) & ":C" & CStr(CurRow - 1) & ")/snapshots/" & CStr(AvgCPUs)
	      aa0 = Cl & CStr(sRow) & ":" & Cl & CStr(CurRow - 1)
	      MyCells(MyRow, 4) = _
		"=SUMPRODUCT(" & aa0 & "," & aa0 & ")/SUM(" & aa0 & ")-C" & CStr(MyRow)
	      MyCells(MyRow, 5) = _
		"=MAX(" & Cl & CStr(sRow) & ":" & Cl & CStr(CurRow - 1) & ")-(C" & CStr(MyRow) & "+D" & CStr(MyRow) & ")"
	      aa0 = Cw & CStr(sRow) & ":" & Cw & CStr(CurRow - 1)
	      MyCells(MyRow, 7) = "=MIN(" & aa0 & ")"
	      MyCells(MyRow, 8) = "=AVERAGE(" & aa0 & ")-G" & CStr(MyRow)
	      MyCells(MyRow, 9) = _
		 "=MAX(" & aa0 & ")-SUM(G" & CStr(MyRow) & ":H" & CStr(MyRow) & ")"
	      aa0 = "J" & CStr(sRow) & ":J" & CStr(CurRow - 1)
	      MyCells(MyRow, 10) = "=AVERAGE(" & aa0 & ")"
	      MyCells(MyRow, 11) = _
		"=IF(SUM(" & aa0 & ")>0,SUMPRODUCT(" & aa0 & "," & aa0 & ")/SUM(" & aa0 & ")-J" & CStr(MyRow) & ",0)"
	      MyCells(MyRow, 12) = _
		 "=MAX(" & aa0 & ")-SUM(J" & CStr(MyRow) & ":K" & CStr(MyRow) & ")"
	      If Cmd = "" Then Exit For
	      sRow = CurRow
	      oldCmd = Cmd
	      UserForm1.Label1.Caption = Cmd
	      UserForm1.ProgressBar1.Value = CurRow
	      UserForm1.Repaint
	   End If
	Next
			     'convert formulas to values so that time stamps can be altered
	MyCells.Range(Cells(2, Cn), Cells(numrows, Cn + 1)).Copy
	MyCells.Range(Cells(2, Cn), Cells(numrows, Cn + 1)).PasteSpecial Paste:=xlPasteValues
	
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
						'produce a graph of CPU by Command
	Set MyCells = Sheet1.Range(Cells(numrows + 2, 3), Cells(MyRow, 5))
	MyCells(1, 1) = "Avg."
	MyCells(1, 2) = "WAvg."
	MyCells(1, 3) = "Max."
	MyCells.NumberFormat = "0.0"
						
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, PlotBy:=xlColumns, CategoryLabels:=0, _
	    SeriesLabels:=1, HasLegend:=True, _
	    Title:="CPU% by command " & Host & "  " & RunDate
	    
	With Chart1.Chart                       'apply customisation
	   .SeriesCollection(1).XValues = "=TOP!R" & CStr(numrows + 3) & "C2:R" & CStr(MyRow) & "C2"
	   Call ApplyStyle(Chart1, 0, 1)
	End With
						'produce a graph of Memory by Command
	Set MyCells = Sheet1.Range(Cells(numrows + 2, 7), Cells(MyRow, 9))
	MyCells(1, 0) = "WSet=>"
	MyCells(1, 1) = "Min."
	MyCells(1, 2) = "Avg."
	MyCells(1, 3) = "Max."
	MyCells.NumberFormat = "0"
						
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, PlotBy:=xlColumns, CategoryLabels:=0, _
	    SeriesLabels:=1, HasLegend:=True, _
	    Title:="Memory by command (MBytes) " & Host & "  " & RunDate
	    
	With Chart1.Chart                       'apply customisation
	   .SeriesCollection(1).XValues = "=TOP!R" & CStr(numrows + 3) & "C2:R" & CStr(MyRow) & "C2"
	   .Axes(xlValue).DisplayUnit = xlThousands
	   .Axes(xlValue).HasDisplayUnitLabel = False
	   Call ApplyStyle(Chart1, 0, 1)
	End With
						'produce a graph of CharIO by Command
	If Not topas Then
	   Set MyCells = Sheet1.Range(Cells(numrows + 2, 10), Cells(MyRow, 12))
	   MyCells(1, 1) = "Avg."
	   MyCells(1, 2) = "WAvg."
	   MyCells(1, 3) = "Max."
	   MyCells.NumberFormat = "0"
						
	   Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop3 + numrows * rH, cWidth, cHeight)
	   Chart1.Chart.ChartWizard Source:=MyCells, PlotBy:=xlColumns, CategoryLabels:=0, _
	      SeriesLabels:=1, HasLegend:=True, _
	      Title:="CharIO by command (bytes/sec) " & Host & "  " & RunDate
	    
	   With Chart1.Chart                       'apply customisation
	      .SeriesCollection(1).XValues = "=TOP!R" & CStr(numrows + 3) & "C2:R" & CStr(MyRow) & "C2"
	      Call ApplyStyle(Chart1, 0, 1)
	   End With
	End If
						'and now produce a graph of CPU by PID
	If Not topas And numrows <= 32000 Then
	   Set MyCells = Sheet1.Range("B1..C" & CStr(numrows))
	   Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop4 + numrows * rH, cWidth, cHeight)
	   Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlXYScatter, _
	      SeriesLabels:=1, CategoryLabels:=1, HasLegend:=False, _
	      Title:="%Processor by PID " & Host & "  " & RunDate
						'apply customisation
	   With Chart1.Chart
	      .Axes(xlCategory).HasMajorGridlines = False
	   End With
	End If
						'scroll the window to the graph
	Sheet1.Range("B2").Select
	ActiveWindow.FreezePanes = True
	Sheet1.Range("A1").Select
	ActiveWindow.ScrollRow = numrows + 2
End Sub
                                       'last mod v3.2.3
Sub PP_UARG(numrows As Long, Sheet1 As Worksheet)
	Dim TopSheet As Worksheet
	Dim MyCells As Range
	
	Sheet1.Columns(4).AutoFit
						'quick fix for Excel formatting problem
	Sheet1.Range("A1") = "Time"
						'if TOP sheet present, add two columns
	If SheetExists("TOP") Then
	   Set TopSheet = Worksheets("TOP")
	   Call GetLastColumn(TopSheet)
	   Set MyCells = TopSheet.Range("A1")
	   MyCells(1, ColNum + 1) = "User"
	   MyCells(1, ColNum + 2) = "Arg"
	   MyCells(2, ColNum + 1) = "=VLOOKUP(B2,UARG!B$2:H$" & CStr(numrows) & ",5,0)"
	   MyCells(2, ColNum + 2) = "=VLOOKUP(B2,UARG!B$2:H$" & CStr(numrows) & ",7,0)"
	   TopSheet.Range(MyCells(2, ColNum + 1), MyCells(TopRows, ColNum + 2)).FillDown
	   TopSheet.Range(MyCells(2, ColNum + 1), MyCells(TopRows, ColNum + 2)).Copy
	   TopSheet.Range(MyCells(2, ColNum + 1), MyCells(TopRows, ColNum + 2)).PasteSpecial Paste:=xlPasteValues
	   End If
						'freeze the headings
	Sheet1.Range("B2").Select
	ActiveWindow.FreezePanes = True
	Sheet1.Range("A1").Select
End Sub
                                       'last mod v3.1.5
Sub PP_WLM(numrows As Long, Sheet1 As Worksheet)
	Dim aa0 As String
	Dim aa1 As String
	Dim Chart1 As ChartObject              'new chart object
	Dim cName As String                    'WLM class name
	Dim Graphdata As Range                 'range for charting
	'Public Host As String                 'Hostname from AAA sheet
	Dim i As Integer                       'column pointer (subclass start)
	Dim j As Integer                       'column pointer (subclass end)
	Dim NewName As String                  'Name of new WLM sheet
	Dim NewSheet As Worksheet
	'Public RunDate As String              'NMON run date from AAA sheet
	Dim SectionName As String
	Dim sName As String                    'WLM subclass name
	    
	SectionName = Sheet1.Name
	Call DelInt(Sheet1, numrows, 1)
	Select Case SectionName
	Case "WLMBIO"
	   Sheet1.Range("A1") = "Block I/O by WLM classes " & Host
	Case "WLMCPU"
	   Sheet1.Range("A1") = "%CPU by WLM classes " & Host
	Case "WLMMEM"
	   Sheet1.Range("A1") = "Memory by WLM classes " & Host
	End Select
					       'handle subclasses
					       'convert subclass names to a std format
	Sheet1.Range("A1:" & LastColumn & "1").Replace What:=",", Replacement:=".", _
	   LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
	i = 7
	Do
	   cName = Sheet1.Cells(1, i) & "."
	   If cName = "." Then Exit Do         'no WLM classes
	   i = i + 1
	   sName = Sheet1.Cells(1, i)
	   If Left(sName, Len(cName)) = cName Then
					       'found start of subclass
	      j = i + 1
	      Do
		 sName = Sheet1.Cells(1, j)
		 If sName = "" Then Exit Do
		 If Left(sName, Len(cName)) <> cName Then Exit Do
		 j = j + 1
	      Loop
					       'found end of subclass
	      j = j - 1
	      If j - i = 0 Then
		 Sheet1.Columns(i).Delete   'no point creating a sheet
	      Else
		 NewName = SectionName & "." & Left(cName, Len(cName) - 1)
					       'create the new sheet
		 Sheets.Add.Name = NewName
		 Set NewSheet = Worksheets(NewName)
		 NewSheet.Move after:=Sheets(Sheets.Count)
		 Sheet1.Columns("A").Copy NewSheet.Range("A1")
		 aa0 = ConvertRef(i - 1) & "1:" & ConvertRef(j - 1) & CStr(numrows)
		 Sheet1.Range(aa0).Copy NewSheet.Range("B1")
		 Sheet1.Range(aa0).Delete
		 Call WLMgraphs(numrows, NewSheet)
		 If NewSheet.Name = "WLMCPU" And SheetExists("LPAR") Then Call WLMPCPU(numrows, NewSheet)
	      End If
	   Else
	   End If
	Loop
	Call WLMgraphs(numrows, Sheet1)
	If Sheet1.Name = "WLMCPU" And SheetExists("LPAR") Then Call WLMPCPU(numrows, Sheet1)
	End Sub
	Sub WLMgraphs(numrows As Long, Sheet1 As Worksheet)        'v3.1.5
	Dim Chart1 As ChartObject
	Dim Graphdata As Range
	Dim SectionName As String
	
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
	SectionName = Sheet1.Name
	Call GetLastColumn(Sheet1)
					       'produce avg/max graph
	Set Graphdata = avgmax(numrows, Sheet1, 0)
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop1 + numrows * rH, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=Graphdata, PlotBy:=xlRows, SeriesLabels:=1, _
	    Title:=Sheet1.Range("A1").Value & "  " & RunDate
					       'apply customisation
	With Chart1.Chart
	     .SeriesCollection(1).XValues = "=" & SectionName & "!R1C2:R1C" & CStr(ColNum)
	     Call ApplyStyle(Chart1, 0, 3)
	End With
					    'produce area graph
	Set Graphdata = Sheet1.Range("A1:" & LastColumn & CStr(numrows))
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, csTop2 + numrows * rH, cWidth, csHeight)
	Chart1.Chart.ChartWizard Source:=Graphdata, Gallery:=xlArea, _
	    Title:=Sheet1.Range("A1").Value & "  " & RunDate, _
	    CategoryLabels:=1, SeriesLabels:=1, HasLegend:=True
					    'apply customisation
	With Chart1.Chart
	     Call ApplyStyle(Chart1, 2, ColNum - 1)
	End With

End Sub
                                       'last mod v3.2.0
Sub PP_ZZZZ(numrows As Long, Sheet1 As Worksheet)
	Dim CurrentRow As Long                  'loop counter
	Dim MyName As String                    'name of target sheet
	Dim numTvals As Long                    'number of time values on target sheet
	Dim NextSheet As Worksheet              'target sheet
	Dim Toffset As Integer                  'Offset for VLOOKUP
	'Public T1 As String                    'first timestamp
	Dim Times As Range                      'pointer to actual times on ZZZZ sheet
	Dim Tvalues As Range                    'area on target sheet
	'Public xToD As String                  'Number format for ToD graphs
	
	Call DelInt(Sheet1, numrows, 0)
	Sheet1.Columns("C").ColumnWidth = 9
	If Sheet1.Cells(1, 3) > "" Then
	   Sheet1.Cells(1, 4) = "=B1+C1"
	   Set Times = Sheet1.Range("D1:D" & CStr(numrows))
	   Times.FillDown
	   Times.Copy
	   Times.PasteSpecial Paste:=xlPasteValues
	   Times.NumberFormat = xToD
	   Toffset = 4
	Else
	   Set Times = Sheet1.Range("B1:B" & CStr(numrows))
	   Toffset = 2
	End If
	T1 = Sheet1.Cells(1, 1)
					    'update the snapshots value on the AAA sheet
	Worksheets("AAA").Range("snapshots").Value = numrows
					    'Then go through each sheet and replace all time values
	UserForm1.ProgressBar1.Value = 0
	UserForm1.ProgressBar1.Max = Worksheets.Count
	UserForm1.Repaint
	For Each NextSheet In Worksheets
	   MyName = NextSheet.Name
	   If MyName = "ZZZZ" Then Exit For
	   With UserForm1
		.Label1.Caption = "Editing Time Values " & MyName
		.Repaint
	   End With
					    'handle TOP/UARG/SUMMARY sheets separately
	   If InStr(1, "SUMMARY#TOP#UARG", MyName) > 0 Then
					    'find how many rows on target sheet
	      For CurrentRow = 2 To 65535
		 If NextSheet.Range("A1").Item(CurrentRow, 1) = "" Then Exit For
		 numTvals = CurrentRow
	      Next
						'build the formulas etc.
	      NextSheet.Range("IV2") = "=VLOOKUP(A2,ZZZZ!A$1:D$" & CStr(numrows) & "," & CStr(Toffset) & ")"
	      If numTvals > 2 Then NextSheet.Range("IV2:IV" & CStr(numTvals)).FillDown
	      NextSheet.Range("IV2:IV" & CStr(numTvals)).Copy
	      NextSheet.Range("IU2").PasteSpecial Paste:=xlPasteValues
	      NextSheet.Range("IV1").EntireColumn.Delete
	      NextSheet.Range("IU2:IU" & CStr(numTvals)).Cut
	      NextSheet.Paste Destination:=NextSheet.Range("A2")
	      NextSheet.Columns("A").NumberFormat = "h:mm:ss"
	   Else
	      If NextSheet.Cells(2, 1).Value = T1 Then
		 Set Tvalues = NextSheet.Range("A2:A" & CStr(numrows + 1))
		 Times.Copy Tvalues
		 Tvalues.NumberFormat = "h:mm:ss"
		 NextSheet.Activate
		 NextSheet.Range("B2").Select
		 ActiveWindow.FreezePanes = True
		 NextSheet.Range("A1").Select
		 ActiveWindow.ScrollRow = numrows + 3
	      End If
	   End If
	   UserForm1.ProgressBar1.Value = UserForm1.ProgressBar1.Value + 1
	Next

End Sub
Function SheetExists(sheetname As String) As Boolean
'returns TRUE if the sheet exists in the active workbook
    SheetExists = False
    On Error GoTo NoSuchSheet
    If Len(Sheets(sheetname).Name) > 0 Then
        SheetExists = True
        Exit Function
    End If
NoSuchSheet:
End Function
                                       'last mod v3.2.6
Sub SVCgraph(numrows As Long, SectionName As String, DoSort As Variant)
	'Public SVCXLIM as Integer              'Lower limit for service time analysis
	Dim aa0 As String                       'temp var
	Dim aa1 As String                       'temp var
	Dim aa2 As String                       'string containing value of SVCXLIM
	Dim bRange As Range                     'range for find method
	'Public Colnum As Integer               'Last column number
	Dim eRange As Range                     'pointer to path numbers
	Dim found As Variant                    'results of find method
	Dim hdisk As String                     'name of hdisk to find in bRange
	'Public Host As String                  'Hostname from AAA sheet
	Dim MyCells As Range                    'temp var
	Dim n As Integer                        'temp var
	Dim NewName As String                   'new name for the SVC sheet
	'Public RunDate As String               'NMON run date from AAA sheet
	Dim Sheet1 As Worksheet                 'pointer to current sheet
	Dim Sheet2 As Worksheet                 'pointer to the SVC sheet
	Dim tmact As String                     'name of the sheet containing device busy stats
	    
	If topas Then Exit Sub
	Set Sheet1 = Worksheets(SectionName)
						    'set up all the strings we need
	n = InStr(1, SectionName, "XFER")
	NewName = SectionName
	Mid(NewName, n, 4) = "SERV"
	tmact = SectionName
	Mid(tmact, n, 4) = "BUSY"
	aa2 = "2>" & CStr(SVCXLIM) & ",+"
	UserForm1.Label1.Caption = NewName
	UserForm1.Repaint
						'create a new sheet for the service times
	Sheet1.Copy after:=Sheets(Sheets.Count)
	Sheets.Item(Sheets.Count).Name = NewName
	Set Sheet2 = Worksheets(NewName)
	Sheet2.Range("A1") = "Est. Service Times(ms) " & Host
						'and now build the formulas
	Set MyCells = Worksheets(NewName).Range("B1:IV1")
	Set bRange = Worksheets(tmact).Range("B1:IV1")
	Set eRange = Worksheets(NewName).Range("B" & CStr(numrows + 6) & ":IU" & CStr(numrows + 6))
						    
	For ColNum = 1 To 255                   'for each hdisk/vpath
	   If MyCells(1, ColNum) = "" Then Exit For
	   hdisk = MyCells(1, ColNum)
						'find out where the busy data is
	   Set found = bRange.Find(hdisk, LookAt:=xlWhole)
	   aa0 = found.AddressLocal(False, True)
	   aa0 = tmact & "!" & Left(aa0, Len(aa0) - 1) & "2"
	   aa1 = SectionName & "!$" & ConvertRef(ColNum)
	   n = eRange(1, ColNum)
	   If n > 1 Then n = n * 10 Else n = 10
	   MyCells(2, ColNum) = "=IF(" & aa1 & aa2 & aa0 & "/" & aa1 & "2*" & CStr(n) & ",0)"
	Next
					    'convert to values to allow XFER sheet to be sorted
	Set MyCells = Sheet2.Range("B2:" & ConvertRef(ColNum - 1) & CStr(numrows))
	MyCells.FillDown
	MyCells.Copy
	MyCells.PasteSpecial Paste:=xlPasteValues
	Application.CutCopyMode = False
	MyCells.NumberFormat = "0.0"
	
	Call DiskGraphs(numrows, NewName, DoSort)
        
End Sub
                                       'last mod v3.2.4
Sub SYS_SUMM()
	Dim aa0 As String                       'temp var
	Dim aa1 As String                       'temp var
	Dim Chart1 As ChartObject               'new chart object
	'Public First As Integer                'First time interval to process
	'Public Last As Long                    'Last time interval to process
	'Public Host As String                  'Hostname from AAA sheet
	Dim MyCells As Range                    'temp var
	Dim numrows As Long
	'Public RunDate As String               'NMON run date from AAA sheet
	Dim SectionName As String               'Name of new sheet
	Dim Sheet1 As Worksheet                 'pointer to System Summary sheet
	Dim shCPU As Worksheet                  'pointer to CPU_ALL sheet
	Dim i As Integer                        'First row for data table
	
	SectionName = "SYS_SUMM"
	UserForm1.Label1.Caption = SectionName
	UserForm1.Repaint
	
	If Not SheetExists("CPU_ALL") Then Exit Sub
	If Not SheetExists("DISK_SUMM") Then Exit Sub
	Set shCPU = Worksheets("CPU_ALL")
	Sheets.Add.Name = SectionName
	Set Sheet1 = Worksheets(SectionName)
	Sheet1.Move Before:=Worksheets("AAA")
	numrows = Worksheets("AAA").Range("snapshots")
						'Produce the graph on SYS_SUMM
	Sheet1.Range("F1").ColumnWidth = 4
						'add top line
	aa0 = CStr(numrows)
	Sheet1.Range("B1").Value = "Samples"
	Sheet1.Range("B1").Font.Bold = True
	Sheet1.Range("C1").Value = Worksheets("AAA").Range("snapshots").Value
	Sheet1.Range("D1").Value = "First"
	Sheet1.Range("D1").Font.Bold = True
	Sheet1.Range("E1").Value = "=INDEX(ZZZZ!A1:B" & aa0 & ",1,2)"
	Sheet1.Range("F1").Value = "Last"
	Sheet1.Range("F1").Font.Bold = True
	Sheet1.Range("G1").Value = "=INDEX(ZZZZ!A1:B" & aa0 & ",snapshots,2)"
	Sheet1.Range("E1:G1").NumberFormat = "h:mm:ss"
						'add I/O and CPU stats (similar to pirat)
	Sheet1.Range("B3").Value = "Total System I/O Statistics"
	Sheet1.Range("B3").Font.Bold = True
	Sheet1.Range("G3").Value = "CPU:"
	Sheet1.Range("G3:L3").Font.Bold = True
	Worksheets("CPU_ALL").Range("B1:F1").Copy Sheet1.Range("H3:L3")
	
	Sheet1.Range("B4").Value = "Avg tps during an interval:"
	Sheet1.Range("E4").Value = "=AVERAGE(DISK_SUMM!D2:D" & aa0 & ")"
	Sheet1.Range("E4:E8").NumberFormat = "#,##0"
	Sheet1.Range("G4").Value = "Avg"
	Sheet1.Range("H4").Value = "=AVERAGE(CPU_ALL!B2:B" & aa0 & ")"
	
	Sheet1.Range("B5").Value = "Max tps during an interval:"
	Sheet1.Range("E5").Value = "=MAX(DISK_SUMM!D2:D" & aa0 & ")"
	Sheet1.Range("G5").Value = "Max"
	Sheet1.Range("H5").Value = "=MAX(CPU_ALL!B2:B" & aa0 & ")"
	
	Sheet1.Range("B6").Value = "Max tps interval time:"
	Sheet1.Range("E6").Value = "=INDEX(DISK_SUMM!A2:A" & aa0 & ",MATCH(E5,DISK_SUMM!D2:D" & aa0 & ",0),1)"
	Sheet1.Range("E6").NumberFormat = "h:mm:ss"
	Sheet1.Range("G6").Value = "Max:Avg"
	Sheet1.Range("H6").Value = "=IF(H4>0,H5/H4,0)"
	Sheet1.Range("H4:L6").FillRight
	Sheet1.Range("H4:L6").NumberFormat = "#,##0.0"
	
	If SheetExists("LPAR") Then
	   Sheet1.Range("M3").Value = "PhysCPU"
	   Sheet1.Range("M4").Value = "=AVERAGE(LPAR!B2:B" & aa0 & ")"
	   Sheet1.Range("M5").Value = "=MAX(LPAR!B2:B" & aa0 & ")"
	   Sheet1.Range("M6").Value = "=M5/M4"
	   Sheet1.Range("M4:M6").NumberFormat = "#,##0.0"
	End If
	
	Sheet1.Range("B7").Value = "Total number of Mbytes read:"
	Sheet1.Range("E7").Value = "=SUM(DISK_SUMM!B2:B" & aa0 & ")*Interval/1000"
	Sheet1.Range("E7").NumberFormat = "#,##0"
	
	Sheet1.Range("B8").Value = "Total number of Mbytes written:"
	Sheet1.Range("E8").Value = "=SUM(DISK_SUMM!C2:C" & aa0 & ")*Interval/1000"
	Sheet1.Range("E8").NumberFormat = "#,##0"
	
	Sheet1.Range("B9").Value = "Read/Write Ratio:"
	Sheet1.Range("E9").Value = "=IF(E8>0,E7/E8,0)"
	Sheet1.Range("E9").NumberFormat = "#,##0.0"
	Sheet1.Range("B1:L29").Copy
	Sheet1.Range("B1:L29").PasteSpecial Paste:=xlPasteValues
	Sheet1.Range("A1").Select
	If Graphs = "LIST" And InStr(1, List, Sheet1.Name) = 0 Then Exit Sub
	If SheetExists("LPAR") Then
	   Set shCPU = Worksheets("LPAR")
	   Set MyCells = shCPU.Range("A1:B" & CStr(numrows + 1))
	Else
	   Set MyCells = Union(shCPU.Range("A1:A" & CStr(numrows + 1)), shCPU.Range("F1:F" & CStr(numrows + 1)))
	End If
	Set Chart1 = Sheet1.ChartObjects.Add(cLeft, cTop, cWidth, cHeight)
	Chart1.Chart.ChartWizard Source:=MyCells, Gallery:=xlLine, Format:=2, _
	   CategoryLabels:=1, SeriesLabels:=1, Title:="System Summary " & Host & "  " & RunDate
	
						'apply customisation
	With Chart1.Chart
	   .SeriesCollection.NewSeries
	   With .SeriesCollection(2)
	     .AxisGroup = 2
	     .MarkerStyle = xlNone
	     .Values = "=DISK_SUMM!R2C4:R" & CStr(numrows + 1) & "C4"
	     .Name = "=DISK_SUMM!R1C4"
	   End With
	   Call ApplyStyle(Chart1, 1, 2)
	   .Axes(xlValue).HasMajorGridlines = False
	   .Axes(xlValue, xlPrimary).HasTitle = True
	   If SheetExists("LPAR") Then
	      .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "#CPUs"
	   Else
	      .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "usr%+sys%"
	      .Axes(xlValue, xlPrimary).MaximumScale = 100
	   End If
	   .Axes(xlValue, xlSecondary).HasTitle = True
	   .Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Disk xfers"
	   .Axes(xlValue, xlSecondary).MinimumScale = 0
	   .SeriesCollection(1).Border.ColorIndex = 25
	   .SeriesCollection(2).Border.ColorIndex = 26
	End With
					       'and move all the text below the graph
	Set MyCells = Chart1.BottomRightCell
	aa1 = CStr(MyCells.Row + 1)
	Sheet1.Range("B3:M9").Cut
	Sheet1.Paste Destination:=Sheet1.Range("B" & aa1)

End Sub
                                       'last mod v3.2.4
Sub TidyUp(CPUList() As String, FirstSheets() As String)
	'Public GotEMC As Variant           'True if either EMC or FAStT present
	'Public GotESS As Variant           'True if EMC/ESS or FAStT present
	'Public NumCPUs As Integer          'number of CPU sheets to move
	'Public Reorder As Variant          'Reorder sheets after analysis (True/False)
	Dim aa0 As String                   'First non-DISK sheet moved
	Dim MyName As String                'name of the current sheet
	Dim n As Integer                    'loop counter
	Dim NextSheet As Worksheet          'temp var
	Dim numdisks As Integer             'number of disk sheets to move
	Dim LastSheet As Worksheet          'anchor for moves
	
	UserForm1.Label1.Caption = "Tidying up ... "
	UserForm1.ProgressBar1.Value = 0
	UserForm1.Repaint
	
	Application.DisplayAlerts = False
	For n = 1 To Application.SheetsInNewWorkbook
	    Worksheets(FirstSheets(n)).Delete
	Next n
	If Not SheetExists("ZZZZ") Then Exit Sub
						'delete empty DISK sheets
	For Each NextSheet In Worksheets
	    If Left$(NextSheet.Name, 4) = "DISK" Then
	       If NextSheet.Range("B1") = "" Then NextSheet.Delete
	    End If
	Next
					       'delete sheets without graphs
	If NoList Then
	   For Each NextSheet In Worksheets
	       If InStr(1, List, NextSheet.Name) = 0 Then NextSheet.Delete
	   Next
	End If
	Application.DisplayAlerts = True
	
	If NoList Or Not Reorder Then Exit Sub
	UserForm1.Label1.Caption = "Re-ordering sheets ... "
	UserForm1.Repaint
	UserForm1.ProgressBar1.Max = Worksheets.Count
	If SheetExists("DISK_SUMM") Then Worksheets("DISK_SUMM").Move Before:=Worksheets(CPUList(1))
	
	If Not GotESS Then
	   Set LastSheet = Worksheets("ZZZZ")
	   For n = 1 To NumCPUs
	      Set NextSheet = Worksheets(CPUList(n))
	      UserForm1.Label1.Caption = "Re-ordering sheets ... " & NextSheet.Name
	      UserForm1.ProgressBar1.Value = UserForm1.ProgressBar1.Value + 1
	      UserForm1.Repaint
	      NextSheet.Move after:=LastSheet
	      Set LastSheet = NextSheet
	   Next n
	Else
						'move ESS/EMC and FILE... sheets
	   Set LastSheet = Worksheets(CPUList(1))
	   For Each NextSheet In Worksheets
	      MyName = NextSheet.Name
	      If MyName = "ZZZZ" Then Exit For
	      If Left$(MyName, 1) > "D" Or Left$(MyName, 2) = "DG" Then
		 UserForm1.Label1.Caption = "Re-ordering sheets ... " & MyName
		 UserForm1.Repaint
		 NextSheet.Move Before:=LastSheet
		 UserForm1.ProgressBar1.Value = UserForm1.ProgressBar1.Value + 1
		 If aa0 = "" Then aa0 = MyName
	      End If
	   Next
	
	   UserForm1.Label1.Caption = "Re-ordering DISK & summary sheets ... "
	   UserForm1.ProgressBar1.Value = UserForm1.ProgressBar1.Max - 5
	   UserForm1.Repaint
	   Set LastSheet = Worksheets(aa0)
	   Worksheets("DISKBUSY").Move Before:=LastSheet
	   If SVCTimes Then Worksheets("DISKSERV").Move Before:=LastSheet
	End If
	If SheetExists("SYS_SUMM") Then Worksheets("SYS_SUMM").Move Before:=Worksheets(1)
	UserForm1.ProgressBar1.Value = UserForm1.ProgressBar1.Max

End Sub
                                       'v3.2.3
Sub WLMPCPU(numrows As Long, Sheet1 As Worksheet)
	Dim aa0 As String
	Dim MyCells As Range
	Dim NewName As String
	Dim NewSheet As Worksheet
					       'create a copy of the CPU sheet
	NewName = Replace(Sheet1.Name, "CPU", "PCPU")
	Sheet1.Copy after:=Sheets(Sheets.Count)
	Sheets.Item(Sheets.Count).Name = NewName
	Set NewSheet = Worksheets(NewName)
	aa0 = NewSheet.Range("A1").Value
	NewSheet.Range("A1") = Replace(aa0, "%", "Physical ")
					       'and convert values to physical CPUs
	Set MyCells = NewSheet.Range("B2:" & LastColumn & CStr(numrows))
	MyCells = "=" & Sheet1.Name & "!B2/100*LPAR!$B2"
	MyCells.Copy
	MyCells.PasteSpecial Paste:=xlPasteValues
	MyCells.NumberFormat = "0.00"
	Call WLMgraphs(numrows, NewSheet)

End Sub

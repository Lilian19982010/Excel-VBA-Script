Attribute VB_Name = "Module2"
Sub Main()
Dim T1 As Date
Dim T2 As Date
Dim T3 As Date
Dim T4 As Date
Dim i As Long
Dim j As Long
Dim rmg As String
Dim FLDR As String
Dim Main As String
Dim TargetFile As String
Dim fileName As String
Dim InputFD As String
Dim Ver As String
Dim fd As Boolean
Dim ProgressForm As New ProgressForm
Dim PBIFDr As String
Dim DashBD As String
Dim chtObj As ChartObject

' To optimize performance, various application settings are disabled:
' - Calculation mode set to manual to prevent unnecessary recalculations.
' - Screen updating turned off to reduce rendering delays.
' - Status bar and events disabled to minimize interference from Excel processes.
' - Page breaks disabled to prevent extra processing overhead.
' However, since the macro takes a long time to execute, ProgressForm.Show vbModeless
' is used to keep the user informed of the current status.
' The form remains non-blocking (vbModeless), allowing execution to continue
' while providing real-time feedback without affecting performance.
ProgressForm.Show vbModeless


Main = "Main"
Ver = "Ver" & ThisWorkbook.Worksheets(Main).Cells(3, 6)
' Ensures that Cell (3, 6) (which corresponds to F6) has the expected status color.
' This color (RGB(3, 252, 111)) serves as an indicator that the previous macro execution completed successfully.
' If the cell color does not match the expected value, a message box warns the user that F6 is not in the correct state.
' The macro execution is halted using "Exit Sub" to prevent further processing until the expected condition is met.
If ThisWorkbook.Worksheets(Main).Cells(3, 6).Interior.Color <> RGB(3, 252, 111) Then
 MsgBox "Cell F6 not in the correct state"
 Exit Sub
End If

T1 = Now
If ThisWorkbook.Worksheets(Main).Cells(3, 6) = 0 Then
 ThisWorkbook.Worksheets(Main).Cells(3, 6) = 1
Else
 ThisWorkbook.Worksheets(Main).Cells(3, 6) = 0
 ThisWorkbook.Worksheets(Main).Cells(3, 6).Interior.ColorIndex = xlNone
End If

' The Power BI Dashboard has duplicate versions, requiring data cleanup before updating.
' The RemoveFiles function clears outdated data sources to ensure a fresh start.
' After removal, the script fetches the latest dataset for the current month,
' ensuring accurate reporting and visualization within the dashboard.
RemoveFiles
CreateNewWorkbook

FLDR = ThisWorkbook.Worksheets(Main).Cells(1, 2)
InputFD = ThisWorkbook.Worksheets(Main).Cells(3, 2)
PBIFDr = ThisWorkbook.Worksheets(Main).Cells(4, 2)
DashBD = "VMM_Dashboard"

ThisWorkbook.Save
ThisWorkbook.SaveAs FLDR & "\" & "VMM Report " & Year(T1) & "-" & Month(T1) & "-" & Day(T1) & Ver & ".xlsb"
ThisWorkbook.Worksheets(Main).Cells(3, 6) = Right(Ver, 1)
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
disableAllPageBreaks
' Remove all chart objects from the DashBoard sheet to reset visual elements.
For Each chtObj In ThisWorkbook.Worksheets(DashBD).ChartObjects
 chtObj.Delete
Next
' Start from row 10 and increment until reaching the row marked "End."
' This helps dynamically locate the last relevant row for processing.
i = 10
Do
 i = i + 1
Loop Until ThisWorkbook.Worksheets(DashBD).Cells(i, 1) = "End"
' Ensure the last identified row is visible for further operations.
ThisWorkbook.Worksheets(DashBD).Rows(i).Hidden = False
ThisWorkbook.Worksheets(DashBD).Activate
' Activate the DashBoard sheet to ensure it is the active worksheet for operations.
ExpandAll DashBD
ThisWorkbook.Worksheets(DashBD).Rows("1:" & CStr(i)).Rows.Ungroup
' Clear any existing row outlines to reset the structure.
ThisWorkbook.Worksheets(DashBD).Rows("1:" & CStr(i)).Rows.ClearOutline

' Clear all relevant data within the defined range (Rows 10 to i, Columns 1 to 46).
ThisWorkbook.Worksheets(DashBD).Range(Cells(10, 1), Cells(i, 46)).Clear
ThisWorkbook.Worksheets(DashBD).Range(Cells(8, 4), Cells(8, 46)).ClearContents
ProgressForm.UpdateProgress 1, "Chart Sheet Cleared"

' Extract supplier purchase order (PO) data from the SupplierPO file.
' Process key performance indicators (KPIs) such as 1st IP On Time, On Time In Full,
' In Full (by Cut-Off), Early Shipment %, and PO Value.
' Perform calculations on extracted data to derive performance metrics.
' Update and publish results to the Vendor Performance Dashboard for visualization.
SupplierPO
ProgressForm.UpdateProgress 2, "Supplier PO Data Catpured!"

ThisWorkbook.Save
' If the version is "Ver0", proceed with copying the shipment sample file
' from the input folder to the Power BI data source folder.
T2 = Now
TargetFile = "Shipment Sample"
fileName = Dir(InputFD & "\" & TargetFile & "*.xls*")
TargetFile = "Shipment Sample.xlsx"
If fileName = "" Then
 GoTo LineXYz
End If
If Ver = "Ver0" Then
 CopyFileToPBI InputFD, fileName, TargetFile
End If

' Extract Ship Sample Approval Data from the Shipment Sample Approval file.
' Process Ship Sample Approval Rate
' Update and publish results to the Vendor Performance Dashboard for visualization.
ShipSampleApproval

' Convert column 14 in "Shipment Sample" to align with Power BI Dashboard format.
' Match values with the reference data in the "Main" worksheet and replace accordingly.
Workbooks.Open PBIFDr & "\Shipment Sample.xlsx"
i = 2
Do
 If Workbooks(TargetFile).Worksheets("Shipment Sample").Cells(i, 14) <> Workbooks(TargetFile).Worksheets("Shipment Sample").Cells(i - 1, 14) Then
  rmg = Workbooks(TargetFile).Worksheets("Shipment Sample").Cells(i, 14)
  j = 3
  fd = False
  Do
   If Workbooks(TargetFile).Worksheets("Shipment Sample").Cells(i, 14) = ThisWorkbook.Worksheets(Main).Cells(j, 31) Then
    Do
     Workbooks(TargetFile).Worksheets("Shipment Sample").Cells(i, 14) = ThisWorkbook.Worksheets(Main).Cells(j, 30)
     i = i + 1
    Loop Until Workbooks(TargetFile).Worksheets("Shipment Sample").Cells(i, 14) <> rmg
    fd = True
   Else
    j = j + 1
   End If
  Loop Until (fd = True)
 End If
Loop Until Workbooks(TargetFile).Worksheets("Shipment Sample").Cells(i, 14) = ""
Workbooks(TargetFile).Close True
ProgressForm.UpdateProgress 3, "Ship Sample Data Captured!"
T3 = Now

' Copy the Factory IQC file to the Power BI data source folder if the version is "Ver0".
' Refine the IQC file to ensure data integrity for publishing in the Power BI Dashboard.
TargetFile = "IQC"
fileName = Dir(InputFD & "\" & TargetFile & "*.xls*")
TargetFile = "IQC.xlsx"
If fileName = "" Then
 GoTo LineXYz
End If
IQCRpt
If Ver = "Ver0" Then
 CopyFileToPBI InputFD, fileName, TargetFile
End If
Workbooks.Open PBIFDr & "\IQC.xlsx"
i = 2
Do
 i = i + 1
Loop Until Workbooks("IQC.xlsx").Worksheets(1).Cells(i, 1) = ""
Workbooks("IQC.xlsx").Worksheets(1).Activate
Workbooks("IQC.xlsx").Worksheets(1).Range(Cells(2, 1), Cells(i - 1, 20)).Sort Key1:=Range(Cells(2, 7), Cells(i - 1, 7)), Order1:=xlAscending, Header:=xlNo
If Workbooks("IQC.xlsx").Worksheets(1).Cells(i, 7) = "-" Then
 i = 2
 Do
  i = i + 1
 Loop Until Workbooks("IQC.xlsx").Worksheets(1).Cells(i, 7) <> "-"
 Workbooks("IQC.xlsx").Worksheets(1).Rows("2:" & CStr(i - 1)).Delete
End If
ProgressForm.UpdateProgress 4, "IQC Data Captured!"

LineXYz:
ThisWorkbook.Save
enableAllPageBreaks
' Create a gauge chart using a combination of Pie Chart and Doughnut Chart.
' The Doughnut Chart serves as the background indicator, displaying the full range.
' The Pie Chart acts as the needle or pointer, highlighting the current value.
' Proper formatting and data setup are essential for correct visual representation.
Chart
ProgressForm.UpdateProgress 5, "Chart Created!"
Name InputFD As InputFD & " - " & Ver
T4 = Now
Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
ThisWorkbook.Save
ProgressForm.Hide
Set ProgressForm = Nothing

MsgBox T1 & " - " & T2 & " - " & T3 & " - " & T4
End Sub

Sub RemoveFiles()
    Dim fso As Object
    Dim filePath As String
    filePath = ThisWorkbook.Worksheets("Main").Cells(4, 2)
    Dim FileP1 As String
    Dim FileP2 As String
    
    Dim FL1 As String
    Dim FL2 As String
    FL1 = "Shipment Sample.xlsx"
    FL2 = "SupplierPO.xlsx"
    
    ' Set the file path
    FileP1 = filePath & FL1
    FileP2 = filePath & FL2
    ' Create the FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the file exists
    If fso.FileExists(FileP1) Then
        ' Delete the file
        fso.DeleteFile FileP1, True ' True indicates to send the file to the recycle bin
    End If
    If fso.FileExists(FileP2) Then
        ' Delete the file
        fso.DeleteFile FileP2, True ' True indicates to send the file to the recycle bin
    End If
    
    ' Clean up
    Set fso = Nothing
End Sub
Sub CreateNewWorkbook()
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add
    Dim filePath As String
    Dim FL1 As String
    Dim FL2 As String
    FL1 = "Shipment Sample.xlsx"
    FL2 = "SupplierPO.xlsx"
    filePath = ThisWorkbook.Worksheets("Main").Cells(4, 2)
    ' Save the new workbook
    newWorkbook.SaveAs filePath & FL1 ' Replace with your desired file path
    
    ' Close the new workbook
    newWorkbook.Close
    
    ' Clean up the object reference
    Set newWorkbook = Nothing
End Sub

Sub ExpandAll(sht As String)
'UpdatebyExtendoffice20181031
    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next
    For i = 1 To 2
        ThisWorkbook.Worksheets(sht).Outline.ShowLevels RowLevels:=i
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next i
    For j = 1 To 2
        ThisWorkbook.Worksheets(sht).Outline.ShowLevels columnLevels:=j
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next j
End Sub

Sub SupplierPO()
'Files Merging
Dim iPO As String
Dim MyFSO As FileSystemObject
Dim MyFolder As Folder
Dim fileName As String
Dim InputFD As String
Dim Main As String
Main = "Main"
Dim FLDR As String
FLDR = ThisWorkbook.Worksheets(Main).Cells(1, 2)
Dim fd As Boolean
Dim fd1 As Boolean
Dim fd2 As Boolean
Dim fd3 As Boolean
Dim fd5 As Boolean
Dim rmg() As Variant
Dim ctr As Long
Dim Rge As Range
Dim SupPrf() As Variant
Dim Transit() As Variant
Dim gRMG() As Variant
Dim MDR1 As Date
Dim MDR2 As Date
Dim MDR3 As Date
Dim P1 As Date
Dim P2 As Date
Dim P3 As Date
Dim LateA As Integer
Dim EarlyA As Integer
Dim Tolerance As Double
Dim IPLog() As Variant
Dim DeleteGP() As String
Dim Y As Long
Dim x As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim n As Long
Dim m As Long
Dim z As Long
Dim FLDT As Date
FLDT = Now
Dim Ys As Integer
Dim Ye As Integer
Dim Ms As Integer
Dim Md As Integer
Dim sti As Long
Dim Edi As Long
Dim Qty As Double
Dim DashBD As String
Dim ArCode() As Variant
Dim Ari As Integer
Dim st As Integer
Dim ed As Integer
Dim OnTime(1 To 20) As Long
Dim ValueS(1 To 9) As Double
Dim Active_Sup() As Variant
ReDim Active_Sup(1 To 4)
DashBD = "VMM_Dashboard"
Dim kk As Integer
Dim FileDate As Date
Dim tmp As Long
Dim S As Long
Dim Ver As String
Ver = "Ver" & ThisWorkbook.Worksheets(Main).Cells(3, 6)
Dim PBIFDr As String
PBIFDr = ThisWorkbook.Worksheets(Main).Cells(4, 2)

MDR1 = DateSerial(ThisWorkbook.Worksheets(Main).Cells(5, 3), ThisWorkbook.Worksheets(Main).Cells(5, 4), 1)
MDR2 = DateSerial(ThisWorkbook.Worksheets(Main).Cells(6, 3), ThisWorkbook.Worksheets(Main).Cells(6, 4), 1)
If ThisWorkbook.Worksheets(Main).Cells(6, 4) = 1 Then
 MDR3 = DateSerial(ThisWorkbook.Worksheets(Main).Cells(6, 3) - 1, 12, 1)
Else
 MDR3 = DateSerial(ThisWorkbook.Worksheets(Main).Cells(6, 3), ThisWorkbook.Worksheets(Main).Cells(6, 4) - 1, 1)
End If
ThisWorkbook.Worksheets(DashBD).Cells(2, 1) = "PO Lines Data : P Date From " & Format(MDR1, "MMM YYYY") & " to " & Format(MDR2, "MMM YYYY")
P3 = DateSerial(ThisWorkbook.Worksheets(Main).Cells(6, 3), ThisWorkbook.Worksheets(Main).Cells(6, 4), 1)

If Month(P3) = 1 Then
 P2 = DateSerial(Year(P3) - 1, 12, 1)
Else
 P2 = DateSerial(Year(P3), Month(P3) - 1, 1)
End If
If Month(P2) = 1 Then
 P1 = DateSerial(Year(P2) - 1, 12, 1)
Else
 P1 = DateSerial(Year(P2), Month(P2) - 1, 1)
End If
LateA = ThisWorkbook.Worksheets(Main).Cells(10, 2)
EarlyA = ThisWorkbook.Worksheets(Main).Cells(12, 2)
Tolerance = ThisWorkbook.Worksheets(Main).Cells(8, 2)
i = 4
k = 0
Do
 i = i + 1
 k = k + 1
Loop Until ThisWorkbook.Worksheets(Main).Cells(i, 8) = ""
ReDim DeleteGP(1 To k)
i = 1
Do
 DeleteGP(i) = ThisWorkbook.Worksheets(Main).Cells(i + 3, 8)
 i = i + 1
Loop Until i > UBound(DeleteGP)
i = 2
ctr = 0
Do
 i = i + 1
 ctr = ctr + 1
Loop Until ThisWorkbook.Worksheets("Suppliers").Cells(i, 1) = ""
ThisWorkbook.Worksheets("Suppliers").Activate
ReDim SupPrf(1 To ctr, 1 To 5)
Set Rge = ThisWorkbook.Worksheets("Suppliers").Range(Cells(1, 1), Cells(i, 14))
SupPrf = Rge.Value

i = 2
ctr = 0
Do
 i = i + 1
 ctr = ctr + 1
Loop Until ThisWorkbook.Worksheets("TransitT").Cells(i, 1) = ""
ThisWorkbook.Worksheets("TransitT").Activate
ReDim Transit(1 To ctr, 1 To 7)
ThisWorkbook.Worksheets("TransitT").Range(Cells(2, 1), Cells(i, 10)).Sort Key1:=Range(Cells(2, 1), Cells(i - 1, 1)), Order1:=xlAscending, Key2:=Range(Cells(2, 2), Cells(i - 1, 2)), Order2:=xlAscending, Key3:=Range(Cells(2, 3), Cells(i - 1, 3)), Order3:=xlAscending, Header:=xlNo
Set Rge = ThisWorkbook.Worksheets("TransitT").Range(Cells(1, 1), Cells(i - 1, 6))
Transit = Rge.Value

IP IPLog
iPO = "SupplierPO"
InputFD = ThisWorkbook.Worksheets(Main).Cells(3, 2)
Ys = ThisWorkbook.Worksheets(Main).Cells(5, 3)
Ye = ThisWorkbook.Worksheets(Main).Cells(6, 3)
Ms = ThisWorkbook.Worksheets(Main).Cells(5, 4)
Md = ThisWorkbook.Worksheets(Main).Cells(6, 4)

If Ms = 1 Then
 Ys = Ys - 1
 Ms = 12
Else
 Ms = Ms - 1
End If

If Md = 12 Then
 Ye = Ye + 1
 Md = 1
Else
 Md = Md + 1
End If

Set MyFSO = New Scripting.FileSystemObject
Set MyFolder = MyFSO.GetFolder(InputFD)

fileName = Dir(InputFD & "\" & iPO & "*.xls*")
Workbooks.Open fileName:=InputFD & "\" & fileName, ReadOnly:=True
FileDate = FileDateTime(InputFD & "\" & fileName)
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
disableAllPageBreaks
 
i = 2
Do
 i = i + 1
Loop Until Workbooks(fileName).Worksheets(1).Cells(i, 1) = ""
Workbooks(fileName).Worksheets(1).Activate
Workbooks(fileName).Worksheets(1).Range(Cells(2, 1), Cells(i - 1, 26)).Sort Key1:=Range(Cells(2, 5), Cells(i - 1, 5)), Order1:=xlAscending, Header:=xlNo

fd = False
j = 2
Do
 'MsgBox Year(Workbooks(Filename).Worksheets(1).Cells(j, 5))
 If (DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(j, 5)), Month(Workbooks(fileName).Worksheets(1).Cells(j, 5)), 1) <= DateSerial(Ys, Ms, 1)) Then
  j = j + 1
 Else
  fd = True
 End If
Loop Until (j <= 1) Or (fd = True)
Workbooks(fileName).Worksheets(1).Rows(CStr(2) & ":" & CStr(j - 1)).Delete
i = 2
Do
 i = i + 1
Loop Until Workbooks(fileName).Worksheets(1).Cells(i, 1) = ""
 
fd = False
j = i
Do
 If (Workbooks(fileName).Worksheets(1).Cells(j, 5) >= DateSerial(Ye, Md, 1)) Or (Workbooks(fileName).Worksheets(1).Cells(j, 5) = "") Then
  j = j - 1
 Else
  fd = True
 End If
Loop Until (j <= 1) Or (fd = True)
 
 
Workbooks(fileName).Worksheets(1).Rows(CStr(j + 1) & ":" & CStr(i - 1)).Delete

Workbooks(fileName).Worksheets(1).Columns(4).Delete
Workbooks(fileName).Worksheets(1).Columns(5).Delete
For i = 1 To 2
 Workbooks(fileName).Worksheets(1).Columns(10).Delete
Next
For i = 1 To 4
 Workbooks(fileName).Worksheets(1).Columns(12).Delete
Next
i = 2
Do
 i = i + 1
Loop Until Workbooks(fileName).Worksheets(1).Cells(i, 1) = ""
 
Workbooks(fileName).Worksheets(1).Range(Cells(2, 1), Cells(i - 1, 26)).Sort Key1:=Range(Cells(2, 10), Cells(i - 1, 10)), Order1:=xlAscending, Header:=xlNo
'remove CPT, CUT, FMM and 0 receipt
x = 2
For Y = 1 To UBound(DeleteGP)
 i = x
 fd = False
 Do
  If Workbooks(fileName).Worksheets(1).Cells(i, 10) = DeleteGP(Y) Then
   m = i
   k = 0
   Do
    k = k + 1
    i = i + 1
   Loop Until (Workbooks(fileName).Worksheets(1).Cells(i, 10) <> DeleteGP(Y)) Or (Workbooks(fileName).Worksheets(1).Cells(i, 1) = "")
   n = i - 1
   Workbooks(fileName).Worksheets(1).Rows(CStr(m) & ":" & CStr(n)).Delete
   x = m
   fd = True
  Else
   i = i + 1
  End If
 Loop Until (Workbooks(fileName).Worksheets(1).Cells(i, 1) = "") Or (fd = True)
Next
'Remove 0 Receipts if P date is one month ealier than the End Range in the Main Page. Which is considered as cancelled, so not be counted.
i = 2
Do
 If Workbooks(fileName).Worksheets(1).Cells(i, 14) = "" Then
  Workbooks(fileName).Worksheets(1).Cells(i, 14) = "TRU"
 End If
 i = i + 1
Loop Until Workbooks(fileName).Worksheets(1).Cells(i, 1) = ""
Workbooks(fileName).Worksheets(1).Range(Cells(2, 1), Cells(i - 1, 26)).Sort Key1:=Range(Cells(2, 8), Cells(i - 1, 8)), Order1:=xlAscending, Key2:=Range(Cells(2, 4), Cells(i - 1, 4)), Order2:=xlAscending, Header:=xlNo
i = 2
fd = False
Do
 If (Workbooks(fileName).Worksheets(1).Cells(i, 8) = 0) And (Workbooks(fileName).Worksheets(1).Cells(i, 4) < MDR3) Then
  m = i
  k = 0
  Do
   k = k + 1
   i = i + 1
  Loop Until (Workbooks(fileName).Worksheets(1).Cells(i, 8) > 0) Or (Workbooks(fileName).Worksheets(1).Cells(i, 1) = "") Or (Workbooks(fileName).Worksheets(1).Cells(i, 4) >= MDR3)
  n = i - 1
  Workbooks(fileName).Worksheets(1).Rows(CStr(m) & ":" & CStr(n)).Delete
  fd = True
 Else
  i = i + 1
 End If
Loop Until (Workbooks(fileName).Worksheets(1).Cells(i, 1) = "") Or (fd = True)
 
i = 2
Do
 i = i + 1
Loop Until Workbooks(fileName).Worksheets(1).Cells(i, 1) = ""
Workbooks(fileName).Worksheets(1).Range(Cells(2, 1), Cells(i - 1, 26)).Sort Key1:=Range(Cells(2, 2), Cells(i - 1, 2)), Order1:=xlAscending, Key2:=Range(Cells(2, 13), Cells(i - 1, 13)), Order2:=xlAscending, Key3:=Range(Cells(2, 14), Cells(i - 1, 14)), Order3:=xlAscending, Header:=xlNo

k = 2
tmp = 2
Do
 j = tmp
 fd5 = False
 Do
  If SupPrf(j, 1) = Workbooks(fileName).Worksheets(1).Cells(k, 2) Then
   tmp = j
   Do
    Workbooks(fileName).Worksheets(1).Cells(k, 16) = SupPrf(j, 2)
    k = k + 1
   Loop Until SupPrf(j, 1) <> Workbooks(fileName).Worksheets(1).Cells(k, 2)
   fd5 = True
  Else
   j = j + 1
  End If
 Loop Until (fd5 = True) Or (j > UBound(SupPrf, 1))
Loop Until Workbooks(fileName).Worksheets(1).Cells(k, 1) = ""

k = 2
tmp = 2
j = tmp
fd = False
Do
 If fd = False Then
LineM:
  j = tmp
 End If
 fd = False
 Do
  If (Transit(j, 1) = Workbooks(fileName).Worksheets(1).Cells(k, 2)) And (Transit(j, 2) = Workbooks(fileName).Worksheets(1).Cells(k, 13)) And (Transit(j, 3) = Workbooks(fileName).Worksheets(1).Cells(k, 14)) Then
   tmp = j
LineZ:
   Workbooks(fileName).Worksheets(1).Cells(k, 15) = Transit(j, 5)
   If Workbooks(fileName).Worksheets(1).Cells(k, 15) <> "" Then
    If (Day(Workbooks(fileName).Worksheets(1).Cells(k, 4)) >= 1) And (Day(Workbooks(fileName).Worksheets(1).Cells(k, 4)) <= Workbooks(fileName).Worksheets(1).Cells(k, 15)) Then
     Select Case Transit(j, 6)
      Case 1
       Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)), 1) 'Determin P Month
       Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
      Case 0
       If Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) = 1 Then
        Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)) - 1, 12, 1)
        Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
       Else
        Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) - 1, 1)
        Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
       End If
      Case 2
       If Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) = 12 Then
        Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 1, 1)
        Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
       Else
        Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 1)
        Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
       End If
      Case 3
       If Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) = 12 Then
        Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 2, 1)
        Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
       Else
        Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 2, 1)
        Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
       End If
     End Select
    Else
     If (Day(Workbooks(fileName).Worksheets(1).Cells(k, 4)) >= Workbooks(fileName).Worksheets(1).Cells(k, 15) + 1) And (Day(Workbooks(fileName).Worksheets(1).Cells(k, 4)) <= Day(DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 0))) Then
      Select Case Transit(j, 6)
       Case 0
        Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)), 1)
        If Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) = 12 Then
         Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 1, Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
        Else
         Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
        End If
       
       Case 1
        If Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) = 12 Then
         Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 1, 1)
         Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 1, Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
        Else
         Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 1)
         Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
        End If
       
       Case 2
        If Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) = 12 Then
         Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 2, 1)
         Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 1, Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
        Else
         Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 2, 1)
         Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
        End If
       
       Case 3
        If Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) = 12 Then
         Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 3, 1)
         Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 1, Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
        Else
         Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 3, 1)
         Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, Workbooks(fileName).Worksheets(1).Cells(k, 15)) 'Cut-off Date
        End If
      End Select
     End If
    End If
   Else
    Workbooks(fileName).Worksheets(1).Cells(k, 17) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)), 1) 'e.g. Sep 1
    Workbooks(fileName).Worksheets(1).Cells(k, 18) = DateSerial(Year(Workbooks(fileName).Worksheets(1).Cells(k, 4)), Month(Workbooks(fileName).Worksheets(1).Cells(k, 4)) + 1, 0) 'Cut off day, e.g. Sep 30
   End If
   fd = True
  Else
   j = j + 1
  End If
 Loop Until (fd = True) Or (j > UBound(Transit, 1))
 If fd = False Then
  Do
   k = k + 1
  Loop Until (Workbooks(fileName).Worksheets(1).Cells(k, 2) <> Workbooks(fileName).Worksheets(1).Cells(k - 1, 2)) Or (Workbooks(fileName).Worksheets(1).Cells(k, 13) <> Workbooks(fileName).Worksheets(1).Cells(k - 1, 13)) Or (Workbooks(fileName).Worksheets(1).Cells(k, 14) <> Workbooks(fileName).Worksheets(1).Cells(k - 1, 14))
  GoTo LineM
 End If
 If (Workbooks(fileName).Worksheets(1).Cells(k, 2) = Workbooks(fileName).Worksheets(1).Cells(k + 1, 2)) And (Workbooks(fileName).Worksheets(1).Cells(k, 13) = Workbooks(fileName).Worksheets(1).Cells(k + 1, 13)) And (Workbooks(fileName).Worksheets(1).Cells(k, 14) = Workbooks(fileName).Worksheets(1).Cells(k + 1, 14)) Then
  k = k + 1
  GoTo LineZ
 Else
  k = k + 1
 End If

Loop Until Workbooks(fileName).Worksheets(1).Cells(k, 1) = ""
'To review the sorting
i = 2
Do
 i = i + 1
Loop Until Workbooks(fileName).Worksheets(1).Cells(i, 1) = ""
Workbooks(fileName).Worksheets(1).Range(Cells(2, 1), Cells(i - 1, 26)).Sort Key1:=Range(Cells(2, 16), Cells(i - 1, 16)), Order1:=xlAscending, Key2:=Range(Cells(2, 2), Cells(i - 1, 2)), Order2:=xlAscending, Key3:=Range(Cells(2, 4), Cells(i - 1, 4)), Order3:=xlAscending, Header:=xlNo
Workbooks(fileName).Worksheets(1).Range(Cells(2, 1), Cells(i - 1, 26)).Sort Key1:=Range(Cells(2, 10), Cells(i - 1, 10)), Order1:=xlAscending, Header:=xlNo
kkk = 4
z = 10
Do
 i = 2
 sti = 0
 Edi = 0
 fd3 = False
 Do
  If Workbooks(fileName).Worksheets(1).Cells(i, 10) = ThisWorkbook.Worksheets(Main).Cells(kkk, 13) Then
   sti = i
   ctr = 1
   Do
    ctr = ctr + 1
    i = i + 1
   Loop Until Workbooks(fileName).Worksheets(1).Cells(i, 10) <> Workbooks(fileName).Worksheets(1).Cells(i - 1, 10)
   Edi = i - 1
   ReDim rmg(1 To ctr - 1, 1 To 22)
   Workbooks(fileName).Worksheets(1).Activate
   Set Rge = Workbooks(fileName).Worksheets(1).Range(Cells(sti, 1), Cells(Edi, 23))
   rmg = Rge.Value
   k = 1
   ctr = 1
   Do
    If k = UBound(rmg, 1) Then
     k = k + 1
     GoTo LineX
    Else
     If (rmg(k, 16) <> rmg(k + 1, 16)) And (k + 1 <= UBound(rmg, 1)) Then
      ctr = ctr + 1
      k = k + 1
     Else
      k = k + 1
     End If
    End If
LineX:
   Loop Until (k > UBound(rmg, 1))
   ReDim gRMG(0 To ctr, 1 To 34)
   ReDim ArCode(1 To ctr, 1 To 20)
   k = 1
   j = 1
   Do
    fd = False
    Do
     If (rmg(k, 17) >= MDR1) And (rmg(k, 17) <= MDR2) Then
      fd = True
     Else
      k = k + 1
     End If
    Loop Until (fd = True) Or (k > UBound(rmg, 1))
    If fd = True Then
     gRMG(j, 1) = rmg(k, 10)
     gRMG(j, 28) = rmg(k, 11)
     gRMG(j, 2) = rmg(k, 16)
     Do
      If (rmg(k, 17) >= MDR1) And (rmg(k, 17) <= MDR2) Then
       Ari = 1
       fd2 = False
       If ArCode(j, Ari) = Empty Then
        If fd2 = False Then
         ArCode(j, Ari) = rmg(k, 2)
        End If
       Else
        Do
         If ArCode(j, Ari) = rmg(k, 2) Then
          fd2 = True
         Else
          Ari = Ari + 1
         End If
        Loop Until (fd2 = True) Or (ArCode(j, Ari) = "")
        If fd2 = False Then
         ArCode(j, Ari) = rmg(k, 2)
        End If
       End If
       gRMG(j, 11) = gRMG(j, 11) + 1
       gRMG(0, 11) = gRMG(0, 11) + 1
       OnTime(4) = OnTime(4) + 1
       gRMG(j, 7) = gRMG(j, 7) + rmg(k, 7) * rmg(k, 9)
       gRMG(0, 7) = gRMG(0, 7) + rmg(k, 7) * rmg(k, 9)
       If rmg(k, 5) <> "" Then
        If rmg(k, 5) - rmg(k, 12) <= rmg(k, 4) + LateA Then
         gRMG(j, 15) = gRMG(j, 15) + 1
         gRMG(0, 15) = gRMG(0, 15) + 1
         rmg(k, 19) = "Yes"
         OnTime(8) = OnTime(8) + 1
        End If
        If rmg(k, 5) - rmg(k, 12) < rmg(k, 4) - EarlyA Then
         gRMG(j, 23) = gRMG(j, 23) + 1
         gRMG(0, 23) = gRMG(0, 23) + 1
         rmg(k, 22) = "Yes"
         OnTime(16) = OnTime(16) + 1
        End If
       End If
      
       If rmg(k, 6) <> "" Then
        If (rmg(k, 6) - rmg(k, 12) <= rmg(k, 4) + LateA) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
         gRMG(j, 19) = gRMG(j, 19) + 1
         gRMG(0, 19) = gRMG(0, 19) + 1
         rmg(k, 20) = "Yes"
         OnTime(12) = OnTime(12) + 1
        End If
        If (rmg(k, 6) - rmg(k, 12) <= rmg(k, 18)) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
         gRMG(j, 27) = gRMG(j, 27) + 1
         gRMG(0, 27) = gRMG(0, 27) + 1
         rmg(k, 21) = "Yes"
         OnTime(20) = OnTime(20) + 1
        End If
        If (rmg(k, 6) - rmg(k, 12) > rmg(k, 18)) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
         n = 1
         fd1 = False
         Qty = 0
         Do
          If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) Then
           m = n
           Do
            Qty = Qty + IPLog(m, 4)
            m = m + 1
           Loop Until (IPLog(m, 2) > rmg(k, 18)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
           If (Qty >= rmg(k, 7) * (1 - Tolerance)) Then
            gRMG(j, 27) = gRMG(j, 27) + 1
            gRMG(0, 27) = gRMG(0, 27) + 1
            rmg(k, 21) = "Yes"
            OnTime(20) = OnTime(20) + 1
           End If
           fd1 = True
          Else
           n = n + 1
          End If
         Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
        End If
       Else
        If rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance) Then
         n = 1
         fd1 = False
         Qty = 0
         Do
          If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) Then
           m = n
           Do
            Qty = Qty + IPLog(m, 4)
            m = m + 1
           Loop Until (IPLog(m, 2) > rmg(k, 18)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
           If (Qty >= rmg(k, 7) * (1 - Tolerance)) Then
            gRMG(j, 27) = gRMG(j, 27) + 1
            gRMG(0, 27) = gRMG(0, 27) + 1
            rmg(k, 21) = "Yes"
            OnTime(20) = OnTime(20) + 1
           End If
           fd1 = True
          Else
           n = n + 1
          End If
         Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
        End If
       End If
       'output to RMG(k, 23)
       If rmg(k, 6) <> "" Then
         If (rmg(k, 6) - rmg(k, 12) < rmg(k, 4) - EarlyA) And (rmg(k, 8) > 0) Then
          rmg(k, 23) = rmg(k, 8) * rmg(k, 9)
         End If
         If (rmg(k, 6) - rmg(k, 12) > rmg(k, 4) - EarlyA) And (rmg(k, 8) > 0) Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) And (IPLog(n, 2) <= rmg(k, 4) - EarlyA) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4) * rmg(k, 9)
             m = m + 1
            Loop Until (IPLog(m, 2) > (rmg(k, 4) - EarlyA)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            rmg(k, 23) = rmg(k, 23) + Qty
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
         End If
       Else
        If rmg(k, 8) > 0 Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) And (IPLog(n, 2) <= rmg(k, 4) - EarlyA) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4) * rmg(k, 9)
             m = m + 1
            Loop Until (IPLog(m, 2) > (rmg(k, 4) - EarlyA)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            rmg(k, 23) = rmg(k, 23) + Qty
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
        End If
       End If
       If rmg(k, 17) = P3 Then
        gRMG(j, 10) = gRMG(j, 10) + 1
        gRMG(0, 10) = gRMG(0, 10) + 1
        OnTime(3) = OnTime(3) + 1
        gRMG(j, 6) = gRMG(j, 6) + rmg(k, 7) * rmg(k, 9)
        gRMG(0, 6) = gRMG(0, 6) + rmg(k, 7) * rmg(k, 9)
        ValueS(3) = ValueS(3) + rmg(k, 7) * rmg(k, 9)
        gRMG(j, 31) = gRMG(j, 31) + rmg(k, 8) * rmg(k, 9)
        gRMG(0, 31) = gRMG(0, 31) + rmg(k, 8) * rmg(k, 9)
        ValueS(6) = ValueS(6) + rmg(k, 8) * rmg(k, 9)
        If rmg(k, 5) <> "" Then
         If rmg(k, 5) - rmg(k, 12) <= rmg(k, 4) + LateA Then
          gRMG(j, 14) = gRMG(j, 14) + 1
          gRMG(0, 14) = gRMG(0, 14) + 1
          OnTime(7) = OnTime(7) + 1
         End If
         If rmg(k, 5) - rmg(k, 12) < rmg(k, 4) - EarlyA Then
          gRMG(j, 22) = gRMG(j, 22) + 1
          gRMG(0, 22) = gRMG(0, 22) + 1
          OnTime(15) = OnTime(15) + 1
         End If
        End If
        If rmg(k, 6) <> "" Then
         If (rmg(k, 6) - rmg(k, 12) <= rmg(k, 4) + LateA) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
          gRMG(j, 18) = gRMG(j, 18) + 1
          gRMG(0, 18) = gRMG(0, 18) + 1
          OnTime(11) = OnTime(11) + 1
         End If
         If (rmg(k, 6) - rmg(k, 12) <= rmg(k, 18)) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
          gRMG(j, 26) = gRMG(j, 26) + 1
          gRMG(0, 26) = gRMG(0, 26) + 1
          OnTime(19) = OnTime(19) + 1
         End If
         If (rmg(k, 6) - rmg(k, 12) < rmg(k, 4) - EarlyA) And (rmg(k, 8) > 0) Then
          gRMG(j, 34) = gRMG(j, 34) + rmg(k, 8) * rmg(k, 9)
          gRMG(0, 34) = gRMG(0, 34) + rmg(k, 8) * rmg(k, 9)
          ValueS(9) = ValueS(9) + rmg(k, 8) * rmg(k, 9)
          
         End If
         If (rmg(k, 6) - rmg(k, 12) > rmg(k, 4) - EarlyA) And (rmg(k, 8) > 0) Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) And (IPLog(n, 2) <= rmg(k, 4) - EarlyA) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4) * rmg(k, 9)
             m = m + 1
            Loop Until (IPLog(m, 2) > (rmg(k, 4) - EarlyA)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            gRMG(j, 34) = gRMG(j, 34) + Qty
            gRMG(0, 34) = gRMG(0, 34) + Qty
            ValueS(9) = ValueS(9) + Qty
            'RMG(k, 23) = RMG(k, 23) + Qty
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
         End If
         If (rmg(k, 6) - rmg(k, 12) > rmg(k, 18)) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4)
             m = m + 1
            Loop Until (IPLog(m, 2) > rmg(k, 18)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            If (Qty >= rmg(k, 7) * (1 - Tolerance)) Then
             gRMG(j, 26) = gRMG(j, 26) + 1
             gRMG(0, 26) = gRMG(0, 26) + 1
             OnTime(19) = OnTime(19) + 1
            End If
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
         End If
        Else
         If rmg(k, 8) > 0 Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) And (IPLog(n, 2) <= rmg(k, 4) - EarlyA) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4) * rmg(k, 9)
             m = m + 1
            Loop Until (IPLog(m, 2) > (rmg(k, 4) - EarlyA)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            gRMG(j, 34) = gRMG(j, 34) + Qty
            gRMG(0, 34) = gRMG(0, 34) + Qty
            ValueS(9) = ValueS(9) + Qty
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
          If (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
           n = 1
           fd1 = False
           Qty = 0
           Do
            If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) Then
             m = n
             Do
              Qty = Qty + IPLog(m, 4)
              m = m + 1
             Loop Until (IPLog(m, 2) > rmg(k, 18)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
             If (Qty >= rmg(k, 7) * (1 - Tolerance)) Then
              gRMG(j, 26) = gRMG(j, 26) + 1
              gRMG(0, 26) = gRMG(0, 26) + 1
              OnTime(19) = OnTime(19) + 1
             End If
             fd1 = True
            Else
             n = n + 1
            End If
           Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
          End If
         End If
        End If
       End If
       If rmg(k, 17) = P2 Then
        gRMG(j, 9) = gRMG(j, 9) + 1
        gRMG(0, 9) = gRMG(0, 9) + 1
        OnTime(2) = OnTime(2) + 1
        gRMG(j, 5) = gRMG(j, 5) + rmg(k, 7) * rmg(k, 9)
        gRMG(0, 5) = gRMG(0, 5) + rmg(k, 7) * rmg(k, 9)
        ValueS(2) = ValueS(2) + rmg(k, 7) * rmg(k, 9)
        gRMG(j, 30) = gRMG(j, 30) + rmg(k, 8) * rmg(k, 9)
        gRMG(0, 30) = gRMG(0, 30) + rmg(k, 8) * rmg(k, 9)
        ValueS(5) = ValueS(5) + rmg(k, 8) * rmg(k, 9)
        If rmg(k, 5) <> "" Then
         If rmg(k, 5) - rmg(k, 12) <= rmg(k, 4) + LateA Then
          gRMG(j, 13) = gRMG(j, 13) + 1
          gRMG(0, 13) = gRMG(0, 13) + 1
          OnTime(6) = OnTime(6) + 1
         End If
        End If
        If rmg(k, 5) <> "" Then
         If rmg(k, 5) - rmg(k, 12) < rmg(k, 4) - EarlyA Then
          gRMG(j, 21) = gRMG(j, 21) + 1
          gRMG(0, 21) = gRMG(0, 21) + 1
          OnTime(14) = OnTime(14) + 1
         End If
        End If
        If rmg(k, 6) <> "" Then
         If rmg(k, 6) - rmg(k, 12) < rmg(k, 4) - EarlyA Then
          gRMG(j, 33) = gRMG(j, 33) + rmg(k, 8) * rmg(k, 9)
          gRMG(0, 33) = gRMG(0, 33) + rmg(k, 8) * rmg(k, 9)
          ValueS(8) = ValueS(8) + rmg(k, 8) * rmg(k, 9)
         End If
         If (rmg(k, 6) - rmg(k, 12) <= rmg(k, 4) + LateA) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
          gRMG(j, 17) = gRMG(j, 17) + 1
          gRMG(0, 17) = gRMG(0, 17) + 1
          OnTime(10) = OnTime(10) + 1
         End If
         If (rmg(k, 6) - rmg(k, 12) <= rmg(k, 18)) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
          gRMG(j, 25) = gRMG(j, 25) + 1
          gRMG(0, 25) = gRMG(0, 25) + 1
          'rmg(k, 21) = "Yes"
          OnTime(18) = OnTime(18) + 1
         End If
         If (rmg(k, 6) - rmg(k, 12) > rmg(k, 4) - EarlyA) And (rmg(k, 8) > 0) Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) And (IPLog(n, 2) <= rmg(k, 4) - EarlyA) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4) * rmg(k, 9)
             m = m + 1
            Loop Until (IPLog(m, 2) > (rmg(k, 4) - EarlyA)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            gRMG(j, 33) = gRMG(j, 33) + Qty
            gRMG(0, 33) = gRMG(0, 33) + Qty
            ValueS(8) = ValueS(8) + Qty
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
         End If
         If (rmg(k, 6) - rmg(k, 12) > rmg(k, 18)) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4)
             m = m + 1
            Loop Until (IPLog(m, 2) > rmg(k, 18)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            If (Qty >= rmg(k, 7) * (1 - Tolerance)) Then
             gRMG(j, 25) = gRMG(j, 25) + 1
             gRMG(0, 25) = gRMG(0, 25) + 1
             OnTime(18) = OnTime(18) + 1
            End If
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
         End If
        Else
         If (rmg(k, 8) > 0) Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) And (IPLog(n, 2) <= rmg(k, 4) - EarlyA) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4) * rmg(k, 9)
             m = m + 1
            Loop Until (IPLog(m, 2) > (rmg(k, 4) - EarlyA)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            gRMG(j, 33) = gRMG(j, 33) + Qty
            gRMG(0, 33) = gRMG(0, 33) + Qty
            ValueS(8) = ValueS(8) + Qty
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
          If (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
           n = 1
           fd1 = False
           Qty = 0
           Do
            If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) Then
             m = n
             Do
              Qty = Qty + IPLog(m, 4)
              m = m + 1
             Loop Until (IPLog(m, 2) > rmg(k, 18)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
             If (Qty >= rmg(k, 7) * (1 - Tolerance)) Then
              gRMG(j, 25) = gRMG(j, 25) + 1
              gRMG(0, 25) = gRMG(0, 25) + 1
              OnTime(18) = OnTime(18) + 1
             End If
             fd1 = True
            Else
             n = n + 1
            End If
           Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
          End If
         End If
        End If
       End If
       If rmg(k, 17) = P1 Then
        gRMG(j, 8) = gRMG(j, 8) + 1
        gRMG(0, 8) = gRMG(0, 8) + 1
        OnTime(1) = OnTime(1) + 1
        gRMG(j, 4) = gRMG(j, 4) + rmg(k, 7) * rmg(k, 9)
        gRMG(0, 4) = gRMG(0, 4) + rmg(k, 7) * rmg(k, 9)
        ValueS(1) = ValueS(1) + rmg(k, 7) * rmg(k, 9)
        gRMG(j, 29) = gRMG(j, 29) + rmg(k, 8) * rmg(k, 9)
        gRMG(0, 29) = gRMG(0, 29) + rmg(k, 8) * rmg(k, 9)
        ValueS(4) = ValueS(4) + rmg(k, 8) * rmg(k, 9)
        If rmg(k, 5) <> "" Then
         If rmg(k, 5) - rmg(k, 12) <= rmg(k, 4) + LateA Then
          gRMG(j, 12) = gRMG(j, 12) + 1
          gRMG(0, 12) = gRMG(0, 12) + 1
          OnTime(5) = OnTime(5) + 1
         End If
        End If
        If rmg(k, 5) <> "" Then
         If rmg(k, 5) - rmg(k, 12) < rmg(k, 4) - EarlyA Then
          gRMG(j, 20) = gRMG(j, 20) + 1
          gRMG(0, 20) = gRMG(0, 20) + 1
          OnTime(13) = OnTime(13) + 1
         End If
        End If
        If rmg(k, 6) <> "" Then
         If (rmg(k, 6) - rmg(k, 12) <= rmg(k, 4) + LateA) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
          gRMG(j, 16) = gRMG(j, 16) + 1
          gRMG(0, 16) = gRMG(0, 16) + 1
          OnTime(9) = OnTime(9) + 1
         End If
         If (rmg(k, 6) - rmg(k, 12) <= rmg(k, 18)) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
          gRMG(j, 24) = gRMG(j, 24) + 1
          gRMG(0, 24) = gRMG(0, 24) + 1
          'rmg(k, 21) = "Yes"
          OnTime(17) = OnTime(17) + 1
         End If
         If rmg(k, 6) - rmg(k, 12) < rmg(k, 4) - EarlyA Then
          gRMG(j, 32) = gRMG(j, 32) + rmg(k, 8) * rmg(k, 9)
          gRMG(0, 32) = gRMG(0, 32) + rmg(k, 8) * rmg(k, 9)
          ValueS(7) = ValueS(7) + rmg(k, 8) * rmg(k, 9)
          'RMG(k, 23) = RMG(k, 8) * RMG(k, 9)
         End If
         If (rmg(k, 6) - rmg(k, 12) >= rmg(k, 4) - EarlyA) And (rmg(k, 8) > 0) Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) And (IPLog(n, 2) <= rmg(k, 4) - EarlyA) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4) * rmg(k, 9)
             m = m + 1
            Loop Until (IPLog(m, 2) > (rmg(k, 4) - EarlyA)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            gRMG(j, 32) = gRMG(j, 32) + Qty
            gRMG(0, 32) = gRMG(0, 32) + Qty
            ValueS(7) = ValueS(7) + Qty
            'RMG(k, 23) = RMG(k, 23) + Qty
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
         End If
         If (rmg(k, 6) - rmg(k, 12) > rmg(k, 18)) And (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4)
             m = m + 1
            Loop Until (IPLog(m, 2) > rmg(k, 18)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            If (Qty >= rmg(k, 7) * (1 - Tolerance)) Then
             gRMG(j, 24) = gRMG(j, 24) + 1
             gRMG(0, 24) = gRMG(0, 24) + 1
             'rmg(k, 21) = "Yes"
             OnTime(17) = OnTime(17) + 1
            End If
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
         End If
        Else
         If rmg(k, 8) > 0 Then
          n = 1
          fd1 = False
          Qty = 0
          Do
           If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) And (IPLog(n, 2) <= rmg(k, 4) - EarlyA) Then
            m = n
            Do
             Qty = Qty + IPLog(m, 4) * rmg(k, 9)
             m = m + 1
            Loop Until (IPLog(m, 2) > (rmg(k, 4) - EarlyA)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
            gRMG(j, 32) = gRMG(j, 32) + Qty
            gRMG(0, 32) = gRMG(0, 32) + Qty
            ValueS(7) = ValueS(7) + Qty
            'RMG(k, 23) = RMG(k, 23) + Qty
            fd1 = True
           Else
            n = n + 1
           End If
          Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
          If (rmg(k, 8) >= rmg(k, 7) * (1 - Tolerance)) Then
           n = 1
           fd1 = False
           Qty = 0
           Do
            If (rmg(k, 1) = IPLog(n, 8)) And (rmg(k, 3) = IPLog(n, 3)) Then
             m = n
             Do
              Qty = Qty + IPLog(m, 4)
              m = m + 1
             Loop Until (IPLog(m, 2) > rmg(k, 18)) Or (rmg(k, 1) <> IPLog(m, 8)) Or (rmg(k, 3) <> IPLog(m, 3))
             If (Qty >= rmg(k, 7) * (1 - Tolerance)) Then
              gRMG(j, 24) = gRMG(j, 24) + 1
              gRMG(0, 24) = gRMG(0, 24) + 1
              'rmg(k, 21) = "Yes"
              OnTime(17) = OnTime(17) + 1
             End If
             fd1 = True
            Else
             n = n + 1
            End If
           Loop Until (fd1 = True) Or (IPLog(n, 1) = "")
          End If
         End If
        End If
       End If
      End If
      k = k + 1
      If k > UBound(rmg, 1) Then
       GoTo LineY
      End If
     Loop Until rmg(k, 16) <> rmg(k - 1, 16)
     j = j + 1
LineY:
    End If
   Loop Until k > UBound(rmg, 1)
   For x = 1 To UBound(gRMG, 1)
    For j = 1 To UBound(ArCode, 2)
     If ArCode(x, j) <> "" Then
      gRMG(x, 3) = gRMG(x, 3) & ArCode(x, j) & ","
     End If
    Next
   Next
   k = 1
   ThisWorkbook.Worksheets(DashBD).Activate
   z = z + 7
  
   ThisWorkbook.Worksheets(DashBD).Cells(z, 1) = gRMG(k, 1)
   ThisWorkbook.Worksheets(DashBD).Cells(z, 2) = gRMG(k, 28)
   ThisWorkbook.Worksheets(DashBD).Range(Cells(z, 1), Cells(z, 2)).Font.Bold = True
   ThisWorkbook.Worksheets(DashBD).Range(Cells(z, 1), Cells(z, 2)).Font.Size = 13
   For kk = 1 To 3
    ThisWorkbook.Worksheets(DashBD).Cells(z - 1, 3 + kk) = gRMG(0, 3 + kk) 'turn to ranking later
    ThisWorkbook.Worksheets(DashBD).Cells(z - 2, 3 + kk) = gRMG(0, 3 + kk)
    ThisWorkbook.Worksheets(DashBD).Cells(z - 2, 3 + kk).NumberFormat = "#,K"
   Next
   For kk = 1 To 3
    ThisWorkbook.Worksheets(DashBD).Cells(z - 6, 28 + kk * 2) = gRMG(0, 3 + kk)
    ThisWorkbook.Worksheets(DashBD).Cells(z - 6, 28 + kk * 2).NumberFormat = "$#,##0"
   Next
  
   ThisWorkbook.Worksheets(DashBD).Cells(z - 7, 30) = "RM PO Val (USD)"
   ThisWorkbook.Worksheets(DashBD).Cells(z - 7, 30).Font.Bold = True
   ThisWorkbook.Worksheets(DashBD).Cells(z - 7, 30).HorizontalAlignment = xlCenter
   ThisWorkbook.Worksheets(DashBD).Range(Cells(z - 7, 30), Cells(z - 7, 35)).Merge
   For kk = 1 To 3
    ThisWorkbook.Worksheets(DashBD).Cells(z - 4, 28 + kk * 2) = gRMG(0, 28 + kk)
    ThisWorkbook.Worksheets(DashBD).Cells(z - 4, 28 + kk * 2).NumberFormat = "$#,##0"
   Next
  
   ThisWorkbook.Worksheets(DashBD).Cells(z - 5, 30) = "RM REC Val (USD)"
   ThisWorkbook.Worksheets(DashBD).Cells(z - 5, 30).HorizontalAlignment = xlCenter
   ThisWorkbook.Worksheets(DashBD).Cells(z - 5, 30).Font.Bold = True
   ThisWorkbook.Worksheets(DashBD).Range(Cells(z - 5, 30), Cells(z - 5, 35)).Merge
   For kk = 1 To 3
    ThisWorkbook.Worksheets(DashBD).Cells(z - 2, 28 + kk * 2) = gRMG(0, 31 + kk)
    ThisWorkbook.Worksheets(DashBD).Cells(z - 2, 28 + kk * 2).NumberFormat = "$#,##0"
   Next
  
   ThisWorkbook.Worksheets(DashBD).Cells(z - 3, 30) = "Early Shipped Val (USD)"
   ThisWorkbook.Worksheets(DashBD).Cells(z - 3, 30).HorizontalAlignment = xlCenter
   ThisWorkbook.Worksheets(DashBD).Cells(z - 3, 30).Font.Bold = True
   ThisWorkbook.Worksheets(DashBD).Range(Cells(z - 3, 30), Cells(z - 3, 35)).Merge
  
   ThisWorkbook.Worksheets(DashBD).Range(Cells(z - 2, 3), Cells(z - 1, 6)).Borders.LineStyle = xlContinuous
   ThisWorkbook.Worksheets(DashBD).Cells(z - 2, 3) = "Value"
   ThisWorkbook.Worksheets(DashBD).Cells(z - 1, 3) = "Rank"
   ThisWorkbook.Worksheets(DashBD).Range(Cells(z - 2, 3), Cells(z - 1, 3)).HorizontalAlignment = xlRight
  
   For kk = 1 To 4
    ThisWorkbook.Worksheets(DashBD).Cells(z - 1, 4 + 5 * kk) = "=200-R[-1]C-R[-2]C"
    ThisWorkbook.Worksheets(DashBD).Cells(z - 2, 4 + 5 * kk) = 1
    ThisWorkbook.Worksheets(DashBD).Cells(z - 3, 4 + 5 * kk) = "=ROUND(R[3]C[3]*100,0)"
    ThisWorkbook.Worksheets(DashBD).Range(Cells(z - 1, 4 + 5 * kk), Cells(z - 3, 4 + 5 * kk)).Font.ColorIndex = 2
   Next
   
   If gRMG(0, 4) = Empty Then
    ThisWorkbook.Worksheets(DashBD).Cells(z, 30) = Empty
   Else
    ThisWorkbook.Worksheets(DashBD).Cells(z, 30) = gRMG(0, 32) / gRMG(0, 4)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 30).NumberFormat = "0%"
   End If
   If gRMG(0, 5) = Empty Then
    ThisWorkbook.Worksheets(DashBD).Cells(z, 32) = Empty
   Else
    ThisWorkbook.Worksheets(DashBD).Cells(z, 32) = gRMG(0, 33) / gRMG(0, 5)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 32).NumberFormat = "0%"
   End If
   If gRMG(0, 6) = Empty Then
    ThisWorkbook.Worksheets(DashBD).Cells(z, 34) = Empty
   Else
    ThisWorkbook.Worksheets(DashBD).Cells(z, 34) = gRMG(0, 34) / gRMG(0, 6)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 34).NumberFormat = "0%"
   End If
   If gRMG(0, 8) = Empty Then
    ThisWorkbook.Worksheets(DashBD).Cells(z, 10) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 15) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 20) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 25) = Empty
   Else
    ThisWorkbook.Worksheets(DashBD).Cells(z, 10) = gRMG(0, 12) / gRMG(0, 8)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 10).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 15) = gRMG(0, 16) / gRMG(0, 8)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 15).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 20) = gRMG(0, 24) / gRMG(0, 8)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 20).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 25) = gRMG(0, 20) / gRMG(0, 8)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 25).NumberFormat = "0%"
   End If
   If gRMG(0, 9) = Empty Then
    ThisWorkbook.Worksheets(DashBD).Cells(z, 11) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 16) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 21) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 26) = Empty
   Else
    ThisWorkbook.Worksheets(DashBD).Cells(z, 11) = gRMG(0, 13) / gRMG(0, 9)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 11).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 16) = gRMG(0, 17) / gRMG(0, 9)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 16).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 21) = gRMG(0, 25) / gRMG(0, 9)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 21).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 26) = gRMG(0, 21) / gRMG(0, 9)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 26).NumberFormat = "0%"
   End If
   If gRMG(0, 10) = Empty Then
    ThisWorkbook.Worksheets(DashBD).Cells(z, 12) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 17) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 22) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 27) = Empty
   Else
    ThisWorkbook.Worksheets(DashBD).Cells(z, 12) = gRMG(0, 14) / gRMG(0, 10)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 12).Font.Bold = True
    ThisWorkbook.Worksheets(DashBD).Cells(z, 12).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 17) = gRMG(0, 18) / gRMG(0, 10)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 17).Font.Bold = True
    ThisWorkbook.Worksheets(DashBD).Cells(z, 17).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 22) = gRMG(0, 26) / gRMG(0, 10)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 22).Font.Bold = True
    ThisWorkbook.Worksheets(DashBD).Cells(z, 22).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 27) = gRMG(0, 22) / gRMG(0, 10)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 27).Font.Bold = True
    ThisWorkbook.Worksheets(DashBD).Cells(z, 27).NumberFormat = "0%"
   End If
   If gRMG(0, 11) = Empty Then
    ThisWorkbook.Worksheets(DashBD).Cells(z, 13) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 18) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 23) = Empty
    ThisWorkbook.Worksheets(DashBD).Cells(z, 28) = Empty
   Else
    ThisWorkbook.Worksheets(DashBD).Cells(z, 13) = gRMG(0, 15) / gRMG(0, 11)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 13).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 18) = gRMG(0, 19) / gRMG(0, 11)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 18).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 23) = gRMG(0, 27) / gRMG(0, 11)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 23).NumberFormat = "0%"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 28) = gRMG(0, 23) / gRMG(0, 11)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 28).NumberFormat = "0%"
   End If
  
   ThisWorkbook.Worksheets(DashBD).Range(Cells(z, 10), Cells(z, 46)).Interior.ColorIndex = 36
   z = z + 1
   ThisWorkbook.Worksheets(DashBD).Range(Cells(z, 1), Cells(z, 46)).Interior.ColorIndex = 49
   ThisWorkbook.Worksheets(DashBD).Range(Cells(z, 1), Cells(z, 46)).Font.ColorIndex = 2
   ThisWorkbook.Worksheets(DashBD).Cells(z, 1) = "RM Group"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 2) = "Supplier Group"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 3) = "Supplier Codes"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 4) = P1
   ThisWorkbook.Worksheets(DashBD).Cells(z, 4).NumberFormat = "mmm-yy"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 5) = P2
   ThisWorkbook.Worksheets(DashBD).Cells(z, 5).NumberFormat = "mmm-yy"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 6) = P3
   ThisWorkbook.Worksheets(DashBD).Cells(z, 6).NumberFormat = "mmm-yy"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 7) = "12 Months"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 8) = "Total Value"
   For kk = 1 To 4
    ThisWorkbook.Worksheets(DashBD).Cells(z, 5 + kk * 5) = P1
    ThisWorkbook.Worksheets(DashBD).Cells(z, 5 + kk * 5).NumberFormat = "mmm-yy"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 6 + kk * 5) = P2
    ThisWorkbook.Worksheets(DashBD).Cells(z, 6 + kk * 5).NumberFormat = "mmm-yy"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 7 + kk * 5) = P3
    ThisWorkbook.Worksheets(DashBD).Cells(z, 7 + kk * 5).NumberFormat = "mmm-yy"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 8 + kk * 5) = "12 Months"
   Next
  
   ThisWorkbook.Worksheets(DashBD).Cells(z, 30).NumberFormat = "mmm-yy"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 30) = P1
   ThisWorkbook.Worksheets(DashBD).Cells(z, 31) = "%"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 32).NumberFormat = "mmm-yy"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 32) = P2
   ThisWorkbook.Worksheets(DashBD).Cells(z, 33) = "%"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 34).NumberFormat = "mmm-yy"
   ThisWorkbook.Worksheets(DashBD).Cells(z, 34) = P3
   ThisWorkbook.Worksheets(DashBD).Cells(z, 35) = "%"
 
   For kk = 1 To 2
    ThisWorkbook.Worksheets(DashBD).Cells(z, 32 + 5 * kk) = P1
    ThisWorkbook.Worksheets(DashBD).Cells(z, 32 + 5 * kk).NumberFormat = "mmm-yy"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 33 + 5 * kk) = P2
    ThisWorkbook.Worksheets(DashBD).Cells(z, 33 + 5 * kk).NumberFormat = "mmm-yy"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 34 + 5 * kk) = P3
    ThisWorkbook.Worksheets(DashBD).Cells(z, 34 + 5 * kk).NumberFormat = "mmm-yy"
    ThisWorkbook.Worksheets(DashBD).Cells(z, 35 + 5 * kk) = "12 Months"
   Next
  
   z = z + 1
   st = z
   Do
    ThisWorkbook.Worksheets(DashBD).Cells(z, 1) = gRMG(k, 1)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 2) = gRMG(k, 2)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 3) = gRMG(k, 3)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 4) = Round(gRMG(k, 4), 2) 'turn to rank Seq later
    ThisWorkbook.Worksheets(DashBD).Cells(z, 5) = Round(gRMG(k, 5), 2) 'turn to rank Seq later
    ThisWorkbook.Worksheets(DashBD).Cells(z, 6) = Round(gRMG(k, 6), 2) 'turn to rank Seq later
    ThisWorkbook.Worksheets(DashBD).Cells(z, 7) = Round(gRMG(k, 7), 2) 'turn to rank Seq later
    ThisWorkbook.Worksheets(DashBD).Cells(z, 8) = Round(gRMG(k, 7), 2)
    ThisWorkbook.Worksheets(DashBD).Cells(z, 8).NumberFormat = "$#,##0;[Red]$#,##0"
   
    If gRMG(k, 8) = Empty Then
     ThisWorkbook.Worksheets(DashBD).Cells(z, 10) = Empty
     ThisWorkbook.Worksheets(DashBD).Cells(z, 15) = Empty
     ThisWorkbook.Worksheets(DashBD).Cells(z, 20) = Empty
     ThisWorkbook.Worksheets(DashBD).Cells(z, 25) = Empty
    Else
     ThisWorkbook.Worksheets(DashBD).Cells(z, 10) = gRMG(k, 12) / gRMG(k, 8)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 15) = gRMG(k, 16) / gRMG(k, 8)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 20) = gRMG(k, 24) / gRMG(k, 8)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 25) = gRMG(k, 20) / gRMG(k, 8)
    End If
   
    If gRMG(k, 9) = Empty Then
     ThisWorkbook.Worksheets(DashBD).Cells(z, 11) = Empty
     ThisWorkbook.Worksheets(DashBD).Cells(z, 16) = Empty
     ThisWorkbook.Worksheets(DashBD).Cells(z, 21) = Empty
     ThisWorkbook.Worksheets(DashBD).Cells(z, 26) = Empty
    Else
     ThisWorkbook.Worksheets(DashBD).Cells(z, 11) = gRMG(k, 13) / gRMG(k, 9)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 16) = gRMG(k, 17) / gRMG(k, 9)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 21) = gRMG(k, 25) / gRMG(k, 9)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 26) = gRMG(k, 21) / gRMG(k, 9)
    End If
   
    If gRMG(k, 10) = Empty Then
     ThisWorkbook.Worksheets(DashBD).Cells(z, 12) = Empty
     ThisWorkbook.Worksheets(DashBD).Cells(z, 17) = Empty
     ThisWorkbook.Worksheets(DashBD).Cells(z, 22) = Empty
     ThisWorkbook.Worksheets(DashBD).Cells(z, 27) = Empty
    Else
     ThisWorkbook.Worksheets(DashBD).Cells(z, 12) = gRMG(k, 14) / gRMG(k, 10)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 17) = gRMG(k, 18) / gRMG(k, 10)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 22) = gRMG(k, 26) / gRMG(k, 10)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 27) = gRMG(k, 22) / gRMG(k, 10)
    End If
   
    If gRMG(k, 11) = Empty Then
     For kk = 1 To 4
      ThisWorkbook.Worksheets(DashBD).Cells(z, 8 + 5 * kk) = Empty
     Next
    Else
     For kk = 1 To 2
      ThisWorkbook.Worksheets(DashBD).Cells(z, 8 + 5 * kk) = gRMG(k, 11 + 4 * kk) / gRMG(k, 11)
     Next
     ThisWorkbook.Worksheets(DashBD).Cells(z, 23) = gRMG(k, 27) / gRMG(k, 11)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 28) = gRMG(k, 23) / gRMG(k, 11)
    End If
  
    For kk = 1 To 3
     ThisWorkbook.Worksheets(DashBD).Cells(z, 28 + 2 * kk) = gRMG(k, 31 + kk)
     ThisWorkbook.Worksheets(DashBD).Cells(z, 28 + 2 * kk).NumberFormat = "$#,##0"
    Next

    For kk = 1 To 3
     If gRMG(k, 3 + kk) = Empty Then
      ThisWorkbook.Worksheets(DashBD).Cells(z, 29 + 2 * kk) = Empty
     Else
      ThisWorkbook.Worksheets(DashBD).Cells(z, 29 + 2 * kk) = gRMG(k, 31 + kk) / gRMG(k, 3 + kk)
      ThisWorkbook.Worksheets(DashBD).Cells(z, 29 + 2 * kk).NumberFormat = "0%"
     End If
    Next
      
    z = z + 1
    k = k + 1
    If k > UBound(gRMG, 1) Then
     Exit Do
    End If
   Loop Until (k > UBound(gRMG, 1)) Or (gRMG(k, 1) = Empty)
  
   ed = z - 1
   For kk = 1 To 4
    ThisWorkbook.Worksheets(DashBD).Range(Cells(st, 5 + 5 * kk), Cells(ed, 8 + 5 * kk)).NumberFormat = "0%"
   Next
  
   Ranking st, ed, DashBD
   iconsets st, ed, 13
   iconsets st, ed, 18
   iconsets st, ed, 23
   iconsets st, ed, 28
   S = sti
   For k = 1 To UBound(rmg, 1)
    Workbooks(fileName).Worksheets(1).Cells(S, 19) = rmg(k, 19)
    Workbooks(fileName).Worksheets(1).Cells(S, 20) = rmg(k, 20)
    Workbooks(fileName).Worksheets(1).Cells(S, 21) = rmg(k, 21)
    Workbooks(fileName).Worksheets(1).Cells(S, 22) = rmg(k, 22)
    Workbooks(fileName).Worksheets(1).Cells(S, 23) = rmg(k, 23)
    S = S + 1
   Next
   Erase gRMG
   Erase rmg
   Erase ArCode

   fd3 = True
  Else
   i = i + 1
  End If
 Loop Until (fd3 = True) Or (Workbooks(fileName).Worksheets(1).Cells(i, 10) = "")
 kkk = kkk + 1
Loop Until ThisWorkbook.Worksheets(Main).Cells(kkk, 13) = ""

ThisWorkbook.Worksheets(DashBD).Cells(z + 1, 1) = "End"
ThisWorkbook.Worksheets(DashBD).Cells(4, 30) = ValueS(1)
ThisWorkbook.Worksheets(DashBD).Cells(4, 32) = ValueS(2)
ThisWorkbook.Worksheets(DashBD).Cells(4, 34) = ValueS(3)
ThisWorkbook.Worksheets(DashBD).Cells(5, 30) = ValueS(4)
ThisWorkbook.Worksheets(DashBD).Cells(5, 32) = ValueS(5)
ThisWorkbook.Worksheets(DashBD).Cells(5, 34) = ValueS(6)
ThisWorkbook.Worksheets(DashBD).Cells(7, 30) = ValueS(7)
ThisWorkbook.Worksheets(DashBD).Cells(7, 32) = ValueS(8)
ThisWorkbook.Worksheets(DashBD).Cells(7, 34) = ValueS(9)

For kk = 1 To 3
 If ValueS(kk) = Empty Then
  ThisWorkbook.Worksheets(DashBD).Cells(8, 28 + 2 * kk) = Empty
 Else
  ThisWorkbook.Worksheets(DashBD).Cells(8, 28 + 2 * kk) = ValueS(6 + kk) / ValueS(kk)
  ThisWorkbook.Worksheets(DashBD).Cells(8, 28 + 2 * kk).NumberFormat = "0%"
 End If
Next

If OnTime(1) = Empty Then
 For kk = 1 To 4
  ThisWorkbook.Worksheets(DashBD).Cells(8, 5 + 5 * kk) = Empty
 Next
Else
 For kk = 1 To 2
  ThisWorkbook.Worksheets(DashBD).Cells(8, 5 + 5 * kk) = OnTime(1 + 4 * kk) / OnTime(1)
 Next
 ThisWorkbook.Worksheets(DashBD).Cells(8, 25) = OnTime(13) / OnTime(1)
 ThisWorkbook.Worksheets(DashBD).Cells(8, 20) = OnTime(17) / OnTime(1)
End If
If OnTime(2) = Empty Then
 For kk = 1 To 4
  ThisWorkbook.Worksheets(DashBD).Cells(8, 6 + 5 * kk) = Empty
 Next
Else
 For kk = 1 To 2
  ThisWorkbook.Worksheets(DashBD).Cells(8, 6 + 5 * kk) = OnTime(2 + 4 * kk) / OnTime(2)
 Next
 ThisWorkbook.Worksheets(DashBD).Cells(8, 26) = OnTime(14) / OnTime(2)
 ThisWorkbook.Worksheets(DashBD).Cells(8, 21) = OnTime(18) / OnTime(2)
End If
If OnTime(3) = Empty Then
 For kk = 1 To 4
  ThisWorkbook.Worksheets(DashBD).Cells(8, 7 + 5 * kk) = Empty
 Next
Else
 For kk = 1 To 2
  ThisWorkbook.Worksheets(DashBD).Cells(8, 7 + 5 * kk) = OnTime(3 + 4 * kk) / OnTime(3)
  ThisWorkbook.Worksheets(DashBD).Cells(8, 7 + 5 * kk).Font.Bold = True
 Next
 ThisWorkbook.Worksheets(DashBD).Cells(8, 27) = OnTime(15) / OnTime(3)
 ThisWorkbook.Worksheets(DashBD).Cells(8, 27).Font.Bold = True
 ThisWorkbook.Worksheets(DashBD).Cells(8, 22) = OnTime(19) / OnTime(3)
 ThisWorkbook.Worksheets(DashBD).Cells(8, 22).Font.Bold = True
End If
If OnTime(4) = Empty Then
 For kk = 1 To 4
  ThisWorkbook.Worksheets(DashBD).Cells(8, 8 + 5 * kk) = Empty
 Next
Else
 For kk = 1 To 2
  ThisWorkbook.Worksheets(DashBD).Cells(8, 8 + 5 * kk) = OnTime(4 + 4 * kk) / OnTime(4)
 Next
 ThisWorkbook.Worksheets(DashBD).Cells(8, 23) = OnTime(20) / OnTime(4)
 ThisWorkbook.Worksheets(DashBD).Cells(8, 28) = OnTime(16) / OnTime(4)
End If

For kk = 1 To 4
 ThisWorkbook.Worksheets(DashBD).Cells(7, 4 + 5 * kk) = "=200-R[-1]C-R[-2]C"
 ThisWorkbook.Worksheets(DashBD).Cells(6, 4 + 5 * kk) = 1
 ThisWorkbook.Worksheets(DashBD).Cells(5, 4 + 5 * kk) = "=ROUND(R[3]C[3]*100,0)"
 ThisWorkbook.Worksheets(DashBD).Range(Cells(5, 4 + 5 * kk), Cells(7, 4 + 5 * kk)).Font.ColorIndex = 2
Next
RMGRanking DashBD
ActiveSuppliers Active_Sup
For i = 1 To 4
 ThisWorkbook.Worksheets(DashBD).Cells(8, i + 3) = Active_Sup(i)
Next
Workbooks(fileName).Worksheets(1).Cells(1, 15) = "Cut off"
Workbooks(fileName).Worksheets(1).Cells(1, 16) = "Supplier Name"
Workbooks(fileName).Worksheets(1).Cells(1, 17) = "P Month"
Workbooks(fileName).Worksheets(1).Cells(1, 18) = "Cut-off End Date"
Workbooks(fileName).Worksheets(1).Cells(1, 19) = "1st IP On Time"
Workbooks(fileName).Worksheets(1).Cells(1, 20) = "On Time In Full"
Workbooks(fileName).Worksheets(1).Cells(1, 21) = "In Full (by Cut-Off)"
Workbooks(fileName).Worksheets(1).Cells(1, 22) = "1st IP Early"
Workbooks(fileName).Worksheets(1).Cells(1, 23) = "Early IP Amt"
i = 2
'Workbooks(Year(FLDT) & "-" & Month(FLDT) & "-" & Day(FLDT) & "-" & Ver &
Do
 If IsEmpty(Workbooks(fileName).Worksheets(1).Cells(i, 23)) = True Then
  Workbooks(fileName).Worksheets(1).Cells(i, 23) = 0
 End If
 i = i + 1
Loop Until Workbooks(fileName).Worksheets(1).Cells(i, 1) = ""

Application.DisplayAlerts = False
Workbooks(fileName).SaveAs FLDR & "\" & Year(FLDT) & "-" & Month(FLDT) & "-" & Day(FLDT) & "-" & Ver & fileName
'PBIFdr
If Ver = "Ver0" Then
 Workbooks(Year(FLDT) & "-" & Month(FLDT) & "-" & Day(FLDT) & "-" & Ver & fileName).SaveAs PBIFDr & "SupplierPO.xlsx"
  Workbooks("SupplierPO.xlsx").Worksheets("SupplierPO").Activate
  i = 2
  Do
   i = i + 1
  Loop Until Workbooks("SupplierPO.xlsx").Worksheets("SupplierPO").Cells(i, 1) = ""
  Workbooks("SupplierPO.xlsx").Worksheets("SupplierPO").Range(Cells(2, 1), Cells(i - 1, 23)).Sort Key1:=Range(Cells(2, 17), Cells(i - 1, 17)), Order1:=xlAscending, Header:=xlNo
  j = 2
  Do
   j = j + 1
  Loop Until (MDR1 > Workbooks("SupplierPO.xlsx").Worksheets("SupplierPO").Cells(j, 17)) Or (MDR2 < Workbooks("SupplierPO.xlsx").Worksheets("SupplierPO").Cells(j, 17))
  If j < i Then
   Workbooks("SupplierPO.xlsx").Worksheets("SupplierPO").Rows(CStr(j) & ":" & CStr(i)).Delete
  End If
 Workbooks("SupplierPO.xlsx").Close True
Else
 Workbooks(Year(FLDT) & "-" & Month(FLDT) & "-" & Day(FLDT) & "-" & Ver & fileName).Close True
End If


Application.DisplayAlerts = True

End Sub

Sub CopyFileToPBI(InputFD As String, fileName As String, TargetFile As String)
Dim fso As Object
Dim sourceFolder As Object
Dim destinationFolder As Object
Dim sourceFile As Object
Dim destinationFile As Object
' Set the source and destination folder paths
Dim sourceFolderPath As String
Dim destinationFolderPath As String
sourceFolderPath = InputFD & "\"
destinationFolderPath = ThisWorkbook.Worksheets("Main").Cells(4, 2)
    
' Set the source and destination file names
Dim sourceFileName As String
Dim destinationFileName As String
sourceFileName = fileName
destinationFileName = TargetFile
    
' Create the FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")
    
' Get the source and destination folder objects
Set sourceFolder = fso.GetFolder(sourceFolderPath)
Set destinationFolder = fso.GetFolder(destinationFolderPath)
    
' Get the source and destination file objects
Set sourceFile = fso.GetFile(sourceFolderPath & sourceFileName)
Set destinationFile = fso.GetFile(destinationFolderPath & destinationFileName)
    
' Check if the destination file already exists, and delete it if necessary
If fso.FileExists(destinationFile.Path) Then
   fso.DeleteFile destinationFile.Path, True ' Delete the destination file (True indicates to send it to the recycle bin)
End If
    
' Copy the source file to the destination folder with a new name
sourceFile.Copy destinationFolder.Path & "\" & destinationFileName, True ' True indicates to overwrite the destination file if it exists
    
    ' Clean up
    Set fso = Nothing
    Set sourceFolder = Nothing
    Set destinationFolder = Nothing
    Set sourceFile = Nothing
    Set destinationFile = Nothing

'Dim i As Long
'Dim j As Long
'Dim PBI As String
'PBI = "\\hktf1pfs.topformbras.com\hktfb\VMM Reports\DataFor VMM BI\"
'Dim IQC As String
'Dim FL2 As String



End Sub

Sub ShipSampleApproval()
Dim iPO As String
Dim MyFSO As FileSystemObject
Dim MyFolder As Folder
Dim fileName As String
Dim InputFD As String
Dim Main As String
Main = "Main"
Dim fd As Boolean
Dim fd1 As Boolean
Dim ctr As Long
Dim Rge As Range
Dim gRMG() As Variant
Dim MDR1 As Date
Dim MDR2 As Date
Dim P1 As Date
Dim P2 As Date
Dim P3 As Date
Dim i As Long
Dim j As Long
Dim k As Long
Dim sti As Long
Dim Edi As Long
Dim DashBD As String
Dim st As Integer
Dim ValueS(1 To 8) As Long
DashBD = "VMM_Dashboard"
Dim ShipAp() As Variant
Dim S As Integer
Dim e As Integer
Dim kk As Integer
Dim Ver As String
Dim RMGs() As Variant

Ver = "Ver" & ThisWorkbook.Worksheets(Main).Cells(3, 6)
Dim PBIFDr As String
PBIFDr = ThisWorkbook.Worksheets(Main).Cells(4, 2)

MDR1 = DateSerial(ThisWorkbook.Worksheets(Main).Cells(5, 3), ThisWorkbook.Worksheets(Main).Cells(5, 4), 1)
MDR2 = DateSerial(ThisWorkbook.Worksheets(Main).Cells(6, 3), ThisWorkbook.Worksheets(Main).Cells(6, 4), 1)
P3 = DateSerial(ThisWorkbook.Worksheets(Main).Cells(6, 3), ThisWorkbook.Worksheets(Main).Cells(6, 4), 1)
i = 3
Do
 i = i + 1
Loop Until ThisWorkbook.Worksheets(Main).Cells(i, 31) = ""
ThisWorkbook.Worksheets(Main).Activate
Set Rge = ThisWorkbook.Worksheets(Main).Range(Cells(3, 31), Cells(i, 32))
RMGs = Rge.Value

If Month(P3) = 1 Then
 P2 = DateSerial(Year(P3) - 1, 12, 1)
Else
 P2 = DateSerial(Year(P3), Month(P3) - 1, 1)
End If
If Month(P2) = 1 Then
 P1 = DateSerial(Year(P2) - 1, 12, 1)
Else
 P1 = DateSerial(Year(P2), Month(P2) - 1, 1)
End If
iPO = "Shipment Sample"

InputFD = ThisWorkbook.Worksheets(Main).Cells(3, 2)

Set MyFSO = New Scripting.FileSystemObject
Set MyFolder = MyFSO.GetFolder(InputFD)

fileName = Dir(InputFD & "\" & iPO & "*.xls*")
Workbooks.Open fileName:=InputFD & "\" & fileName, ReadOnly:=True
ShipSampleTrim fileName, MDR1, MDR2

i = 2
sti = 0
Edi = 0
Do
 If Workbooks(fileName).Worksheets(1).Cells(i, 7) <> Workbooks(fileName).Worksheets(1).Cells(i - 1, 7) Then
  ctr = 0
  sti = i
  Do
   ctr = ctr + 1
   i = i + 1
  Loop Until Workbooks(fileName).Worksheets(1).Cells(i, 7) <> Workbooks(fileName).Worksheets(1).Cells(i - 1, 7)
  Edi = i - 1
 Else
  i = i + 1
 End If
    
 ReDim ShipAp(1 To ctr, 1 To 8)
 Workbooks(fileName).Worksheets(1).Activate
 Set Rge = Workbooks(fileName).Worksheets(1).Range(Cells(sti, 1), Cells(Edi, 8))
 ShipAp = Rge.Value
 j = 10
 fd = False
 Do
  If (ShipAp(1, 7) = ThisWorkbook.Worksheets(DashBD).Cells(j, 2)) And ("RM Group" = ThisWorkbook.Worksheets(DashBD).Cells(j + 1, 1)) Then
   k = j + 2
   st = k
   ctr = 0
   Do
    ctr = ctr + 1
    k = k + 1
   Loop Until ThisWorkbook.Worksheets(DashBD).Cells(k, 1) = ""
   ReDim gRMG(0 To ctr, 1 To 10)
   k = j + 2
   S = k
   For kk = 1 To UBound(gRMG, 1)
    gRMG(kk, 1) = k
    gRMG(kk, 2) = ThisWorkbook.Worksheets(DashBD).Cells(k, 2)
    k = k + 1
   Next
   e = k - 1
   iconsets S, e, 45
   k = 1
   Do
    kk = 1
    fd1 = False
    Do
     If ShipAp(k, 8) = gRMG(kk, 2) Then
      Do
       If (CInt(ShipAp(k, 6)) = 1) Or (CInt(ShipAp(k, 6)) = 2) Then
        Select Case ShipAp(k, 3) & "-" & ShipAp(k, 4)
         Case Year(P1) & "-" & Month(P1)
          gRMG(kk, 3) = gRMG(kk, 3) + 1
          gRMG(0, 3) = gRMG(0, 3) + 1
          ValueS(1) = ValueS(1) + 1
         Case Year(P2) & "-" & Month(P2)
          gRMG(kk, 4) = gRMG(kk, 4) + 1
          gRMG(0, 4) = gRMG(0, 4) + 1
          ValueS(2) = ValueS(2) + 1
         Case Year(P3) & "-" & Month(P3)
          gRMG(kk, 5) = gRMG(kk, 5) + 1
          gRMG(0, 5) = gRMG(0, 5) + 1
          ValueS(3) = ValueS(3) + 1
        End Select
        gRMG(kk, 6) = gRMG(kk, 6) + 1
        gRMG(0, 6) = gRMG(0, 6) + 1
        ValueS(4) = ValueS(4) + 1
       Else
        Select Case ShipAp(k, 3) & "-" & ShipAp(k, 4)
         Case Year(P1) & "-" & Month(P1)
          gRMG(kk, 7) = gRMG(kk, 7) + 1
          gRMG(0, 7) = gRMG(0, 7) + 1
          ValueS(5) = ValueS(5) + 1
         Case Year(P2) & "-" & Month(P2)
          gRMG(kk, 8) = gRMG(kk, 8) + 1
          gRMG(0, 8) = gRMG(0, 8) + 1
          ValueS(6) = ValueS(6) + 1
         Case Year(P3) & "-" & Month(P3)
          gRMG(kk, 9) = gRMG(kk, 9) + 1
          gRMG(0, 9) = gRMG(0, 9) + 1
          ValueS(7) = ValueS(7) + 1
        End Select
        gRMG(kk, 10) = gRMG(kk, 10) + 1
        gRMG(0, 10) = gRMG(0, 10) + 1
        ValueS(8) = ValueS(8) + 1
       End If
       If k < UBound(ShipAp, 1) Then
        k = k + 1
       Else
        k = k + 1
        GoTo LineX
       End If
      Loop Until (ShipAp(k, 8) <> ShipAp(k - 1, 8)) Or (k > UBound(ShipAp, 1))
      fd1 = True
     Else
      kk = kk + 1
     End If
    Loop Until (fd1 = True) Or (kk > UBound(gRMG, 1))
LineX:
   Loop Until k > UBound(ShipAp, 1)
   ThisWorkbook.Worksheets(DashBD).Activate
   ThisWorkbook.Worksheets(DashBD).Cells(st - 3, 41) = "=200-R[-1]C-R[-2]C"
   ThisWorkbook.Worksheets(DashBD).Cells(st - 4, 41) = 1
   ThisWorkbook.Worksheets(DashBD).Cells(st - 5, 41) = "=ROUND(R[3]C[3]*100,0)"
   ThisWorkbook.Worksheets(DashBD).Range(Cells(st - 5, 41), Cells(st - 3, 41)).Font.ColorIndex = 2
   If (gRMG(0, 3) + gRMG(0, 7)) > 0 Then
    ThisWorkbook.Worksheets(DashBD).Cells(st - 2, 42) = gRMG(0, 3) / (gRMG(0, 3) + gRMG(0, 7))
    ThisWorkbook.Worksheets(DashBD).Cells(st - 2, 42).NumberFormat = "0%"
   End If
   If (gRMG(0, 4) + gRMG(0, 8)) > 0 Then
    ThisWorkbook.Worksheets(DashBD).Cells(st - 2, 43) = gRMG(0, 4) / (gRMG(0, 4) + gRMG(0, 8))
    ThisWorkbook.Worksheets(DashBD).Cells(st - 2, 43).NumberFormat = "0%"
   End If
   If (gRMG(0, 5) + gRMG(0, 9)) > 0 Then
    ThisWorkbook.Worksheets(DashBD).Cells(st - 2, 44) = gRMG(0, 5) / (gRMG(0, 5) + gRMG(0, 9))
    ThisWorkbook.Worksheets(DashBD).Cells(st - 2, 44).NumberFormat = "0%"
   End If
   If (gRMG(0, 6) + gRMG(0, 10)) > 0 Then
    ThisWorkbook.Worksheets(DashBD).Cells(st - 2, 45) = gRMG(0, 6) / (gRMG(0, 6) + gRMG(0, 10))
    ThisWorkbook.Worksheets(DashBD).Cells(st - 2, 45).NumberFormat = "0%"
   End If
   kk = 1
   Do
    If (gRMG(kk, 3) + gRMG(kk, 7)) > 0 Then
     ThisWorkbook.Worksheets(DashBD).Cells(gRMG(kk, 1), 42) = gRMG(kk, 3) / (gRMG(kk, 3) + gRMG(kk, 7))
     ThisWorkbook.Worksheets(DashBD).Cells(gRMG(kk, 1), 42).NumberFormat = "0%"
    End If
    If (gRMG(kk, 4) + gRMG(kk, 8)) > 0 Then
     ThisWorkbook.Worksheets(DashBD).Cells(gRMG(kk, 1), 43) = gRMG(kk, 4) / (gRMG(kk, 4) + gRMG(kk, 8))
     ThisWorkbook.Worksheets(DashBD).Cells(gRMG(kk, 1), 43).NumberFormat = "0%"
    End If
    If (gRMG(kk, 5) + gRMG(kk, 9)) > 0 Then
     ThisWorkbook.Worksheets(DashBD).Cells(gRMG(kk, 1), 44) = gRMG(kk, 5) / (gRMG(kk, 5) + gRMG(kk, 9))
     ThisWorkbook.Worksheets(DashBD).Cells(gRMG(kk, 1), 44).NumberFormat = "0%"
    End If
    If (gRMG(kk, 6) + gRMG(kk, 10)) > 0 Then
     ThisWorkbook.Worksheets(DashBD).Cells(gRMG(kk, 1), 45) = gRMG(kk, 6) / (gRMG(kk, 6) + gRMG(kk, 10))
     ThisWorkbook.Worksheets(DashBD).Cells(gRMG(kk, 1), 45).NumberFormat = "0%"
    End If
    kk = kk + 1
   Loop Until kk > UBound(gRMG, 1)
   Erase gRMG
   Erase ShipAp
   fd = True
  Else
   j = j + 1
  End If
 Loop Until (fd = True) Or (ThisWorkbook.Worksheets(DashBD).Cells(j, 1) = "End")
  
 If (ValueS(1) + ValueS(5)) > 0 Then
  ThisWorkbook.Worksheets(DashBD).Cells(8, 42) = ValueS(1) / (ValueS(1) + ValueS(5))
  ThisWorkbook.Worksheets(DashBD).Cells(8, 42).NumberFormat = "0%"
 End If
 If (ValueS(2) + ValueS(6)) > 0 Then
  ThisWorkbook.Worksheets(DashBD).Cells(8, 43) = ValueS(2) / (ValueS(2) + ValueS(6))
  ThisWorkbook.Worksheets(DashBD).Cells(8, 43).NumberFormat = "0%"
 End If
 If (ValueS(3) + ValueS(7)) > 0 Then
  ThisWorkbook.Worksheets(DashBD).Cells(8, 44) = ValueS(3) / (ValueS(3) + ValueS(7))
  ThisWorkbook.Worksheets(DashBD).Cells(8, 44).NumberFormat = "0%"
 End If
 If (ValueS(4) + ValueS(8)) > 0 Then
  ThisWorkbook.Worksheets(DashBD).Cells(8, 45) = ValueS(4) / (ValueS(4) + ValueS(8))
  ThisWorkbook.Worksheets(DashBD).Cells(8, 45).NumberFormat = "0%"
 End If
 ThisWorkbook.Worksheets(DashBD).Cells(7, 41) = "=200-R[-1]C-R[-2]C"
 ThisWorkbook.Worksheets(DashBD).Cells(6, 41) = 1
 ThisWorkbook.Worksheets(DashBD).Cells(5, 41) = "=ROUND(R[3]C[3]*100,0)"
 ThisWorkbook.Worksheets(DashBD).Activate
 ThisWorkbook.Worksheets(DashBD).Range(Cells(5, 41), Cells(7, 41)).Font.ColorIndex = 2
   
Loop Until Workbooks(fileName).Worksheets(1).Cells(i, 1) = ""


fileName = Dir()
 
End Sub

Public Sub disableAllPageBreaks()

  Dim ws As Worksheet

  For Each ws In ThisWorkbook.Worksheets
    ws.DisplayPageBreaks = False
  Next ws

End Sub
Public Sub enableAllPageBreaks()

  Dim ws As Worksheet

  For Each ws In ThisWorkbook.Worksheets
    ws.DisplayPageBreaks = True
  Next ws

End Sub


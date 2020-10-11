Sub Weekly_API_Suppliers_Gateways_LW_Report()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Call Ol
Call BackData
Call FormattingTotals
Call FormattingLastTotal
Call Report_Outlook
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub Ol()
'-----------------------------------------------------------Outlook Macro Start-------------------------------------------------------------------------------------------
    Dim gappOutlook As Object
Dim ns As Namespace
Dim inbox As MAPIFolder
Dim subfolder As MAPIFolder
Dim item As Object
Dim atmt As Attachment
Dim filename As String
Dim I As Integer
Dim varResponse As VbMsgBoxResult
Set ns = GetNamespace("MAPI")
Set gappOutlook = CreateObject("Outlook.Application")
Set inbox = gappOutlook.Session.Folders("Commercial Commercial")
Set subfolder = inbox.Folders("Inbox")
I = 0
' Check subfolder for messages and exit of none found
If subfolder.Items.Count = 0 Then
    MsgBox "There are no messages in the Subm from Arch folder.", vbInformation, _
           "Nothing Found"
    Exit Sub
End If
For Each item In subfolder.Items
If TypeName(item) = "MailItem" Then 'Change the below string to      WorksheetFunction.WeekNum((item.ReceivedTime)) = WorksheetFunction.WeekNum((Now))-1 _     if for any reason you are running the macro the week after(-1 means one week after, -2 would mean two weeks after, anyway it would be very unlikely,unless you need an old report)
If WorksheetFunction.WeekNum((item.ReceivedTime)) = WorksheetFunction.WeekNum((Now)) _
Then
    For Each atmt In item.Attachments
' Check filename of each attachment and save if it has "xlsx" extension
        If atmt.filename = "API Suppliers-Gateways LWBookings Back Data.xlsx" _
        Then
        With item
.UnRead = False
End With
        ' This path must exist! Change folder name as necessary.
            filename = "***\API Suppliers\" & _
                atmt.filename
            atmt.SaveAsFile filename
            I = I + 1
        End If
    Next atmt
    End If
    End If
Next item

' Clear memory
OlAtmt1stMonth_exit:
Set atmt = Nothing
Set item = Nothing
Set ns = Nothing

'-----------------------------------------------------------Outlook Macro End-------------------------------------------------------------------------------------------
End Sub


Sub BackData()
Dim WB As Workbook
Dim wb2 As Workbook
Dim wb3 As Worksheet
Dim wb4 As Workbook
Dim ws As Worksheet
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim ws22 As Worksheet
Dim LR As Long
Application.ScreenUpdating = False
Set WB = Workbooks.Open(filename:="***\API Suppliers-Gateways LWBookings Back Data.xlsx")
Set wb2 = Workbooks.Open(filename:="***\API Suppliers-Gateways LWBookings Template.xlsx")
Set wb4 = Workbooks.Open(filename:="***\API Suppliers-Gateways LWBookings Report.xlsx")
Set ws = WB.Sheets("API Only_1")
Set ws1 = WB.Sheets("Total_2")
Set ws2 = wb2.Sheets("Back Data")
Set ws22 = wb2.Sheets("Back Data2")
Set ws3 = wb2.Sheets("Template")


With ws
LR = .Range("A4:X1000000").Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
ws2.Range("A4", "X" & LR) = .Range("A4", "X" & LR).Value
.Range("A4", "X" & LR).Copy
ws2.Range("A4").PasteSpecial xlPasteFormats
ws2.Range("A4", "B" & LR).Copy
ws3.Range("B3").PasteSpecial xlPasteValues
ws2.Range("A4", "B" & LR).Copy
ws3.Range("B3").PasteSpecial xlPasteFormats
ws3.Range("D3", "R" & LR).FillDown
ws3.Rows(LR).EntireRow.Delete
End With

With ws1
LR = .Range("A4:B1000000").Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
.Range("A4", "B" & LR).Copy
ws22.Range("A4", "B" & LR).PasteSpecial xlPasteValues
ws22.Range("C4", "D" & LR).FillDown
End With


Application.CutCopyMode = False
WB.Close

Kill ("***\API Suppliers-Gateways LWBookings Back Data.xlsx")

Application.ScreenUpdating = True
    End Sub

Sub FormattingTotals()
Dim WB As Workbook
Dim ba As Worksheet
Dim te As Worksheet
Dim co As Worksheet
Dim fo As Worksheet
Application.ScreenUpdating = False

Set WB = Workbooks.Open(filename:="***\API Suppliers-Gateways LWBookings Template.xlsx")
Set ba = WB.Sheets("Back Data")
Set te = WB.Sheets("Template")
Set fo = WB.Sheets("Formatting")
With WB.Sheets("Back Data").Range("B:B")
     Set C = .Find("Total", LookIn:=xlValues)
          If Not C Is Nothing Then
        firstAddress = C.Address
        Do
            fo.Range("B2:G2").Copy
            te.Cells(C.Row - 1, C.Column + 1).PasteSpecial xlPasteFormats
            Set C = .FindNext(C)

        If C Is Nothing Then
            GoTo DoneFinding
        End If
        Loop While C.Address <> firstAddress
      End If
DoneFinding:
End With
Application.ScreenUpdating = True
End Sub

Sub FormattingLastTotal()
Dim WB As Workbook
Dim ba As Worksheet
Dim te As Worksheet
Dim co As Worksheet
Dim fo As Worksheet
Application.ScreenUpdating = False

Set WB = Workbooks.Open(filename:="***\API Suppliers-Gateways LWBookings Template.xlsx")
Set ba = WB.Sheets("Back Data")
Set te = WB.Sheets("Template")
Set fo = WB.Sheets("Formatting")

With WB.Sheets("Back Data").Range("A:B")
Set C = .Find("Total", LookIn:=xlValues, searchdirection:=xlPrevious)
If Not C Is Nothing Then
        firstAddress = C.Address
        Do
            fo.Range("A3:B3").Copy
            te.Cells(C.Row - 1, C.Column + 1).PasteSpecial xlPasteFormats
            fo.Range("C3:G3").Copy
            te.Cells(C.Row - 1, C.Column + 3).PasteSpecial xlPasteFormats

            Set C = .FindNext(C)

        If C Is Nothing Then
            GoTo DoneFinding
        End If
        Loop While C.Address = firstAddress
      End If
DoneFinding:
End With
Application.ScreenUpdating = True
End Sub

Sub Report_Outlook()

' NEWPATH is THIS year's path
NEWPATH = "***\Weekly Reports " & Format(DateAdd("d", 0, Now), "yyyy")
' OLDPATH is the previous year's path,will be useful when in January you want the Previous month(December) report to be saved in previous year's path, and not in this year's path
OLDPATH = "***\Weekly Reports " & Format(DateAdd("w", -1, Now), "yyyy")
' If the folder doesn't exist, then create it, useful again in February when running the January Report, it will create then new year's Folder and save it in there
If Dir(NEWPATH, vbDirectory) = "" _
Then MkDir NEWPATH

' SAVE report in the Right PATH
If Format(DateAdd("w", 0, Now), "yyyy") = Format(DateAdd("w", -1, Now), "yyyy") _
Then RPATH = NEWPATH & "\API Suppliers-Gateways LWBookings Report W.E. " & Format(DateAdd("d", -Weekday(Now) + 1, Now), "dd.mm.yy") & ".xlsx"
If Not Format(DateAdd("w", 0, Now), "yyyy") = Format(DateAdd("w", -1, Now), "yyyy") _
Then RPATH = OLDPATH & "\API Suppliers-Gateways LWBookings Report W.E. " & Format(DateAdd("d", -Weekday(Now) + 1, Now), "dd.mm.yy") & ".xlsx"

Dim WB As Workbook
Dim wb2 As Workbook
Dim ws As Worksheet
Dim ws2 As Worksheet


Set WB = Workbooks.Open(filename:="***\API Suppliers-Gateways LWBookings Template.xlsx")
Set wb2 = Workbooks.Open(filename:="***\API Suppliers-Gateways LWBookings Report.xlsx")
Set ws = WB.Sheets("Template")
Set ws2 = wb2.Sheets("Report")

ws.Range("B3:R1000000").Copy
ws2.Range("B3:R1000000").PasteSpecial xlPasteValues
ws.Range("B3:R1000000").Copy
ws2.Range("B3:R50000").PasteSpecial xlPasteFormats

WB.Sheets("Back Data").Range("A4:C10000").ClearContents
WB.Sheets("Back Data").Range("A4:C10000").UnMerge
WB.Sheets("Back Data").Range("A4:C10000").Style = "Normal"

ws.Range("D4:R1000000").ClearContents

WB.Sheets("Back Data2").Range("A4:B1000000").ClearContents
WB.Sheets("Back Data2").Range("A4:B1000000").UnMerge
WB.Sheets("Back Data2").Range("A4:B1000000").Style = "Normal"
WB.Sheets("Back Data2").Range("C5:D1000000").ClearContents

ws.Range("B3:C10000").ClearContents
ws.Range("D4:D10000").ClearContents
ws.Range("B3:C10000").UnMerge
ws.Range("B3:C10000").Style = "Normal"
ws.Range("D4:D10000").Style = "Normal"
ws.Range("D4:R1000000").ClearContents
ws.Range("D4:R1000000").Style = "Normal"


Application.CutCopyMode = False

WB.Close

wb2.Sheets("Report").Columns(5).EntireColumn.Delete
wb2.Sheets("Report").Columns(5).EntireColumn.Delete
wb2.Sheets("Report").Columns(7).EntireColumn.Delete
wb2.Sheets("Report").Columns(7).EntireColumn.Delete
wb2.Sheets("Report").Range("A1").Copy
wb2.Sheets("Report").Range("A1").PasteSpecial xlPasteValues
wb2.Sheets("Report").Columns("D:F").HorizontalAlignment = xlCenter
wb2.Sheets("Report").Columns("G:G").ClearContents
wb2.SaveAs filename:=(RPATH)


 Dim olapp As Object
    Dim olmail As Object
    Dim olsubject As String

    Application.ScreenUpdating = False

    Set olapp = CreateObject("Outlook.Application")
    Set olmail = olapp.createitem(olmailitem)

    olsubject = "API Suppliers-Gateways LWBookings Report WE " & Format(Now - (iWeekday - 1), "DD.MM.YY") & ".xlsx"

    With olmail
        .display
    End With

    With olmail
        .To = "***@***"
        .CC = "***@***"
        .BCC = ""
        .Subject = olsubject
        .HTMLBody = "Please find attached the Weekly Report for Last Week's Bookings (API Suppliers-Gateways)" & .HTMLBody
        .Attachments.Add (ActiveWorkbook.FullName)
        '.Attachments.Add ("C:\test.txt") ' add other file
'        .Send   'or use .Display
        .display
    End With

    Set olmail = Nothing
    Set olapp = Nothing

wb2.Close
Application.ScreenUpdating = True

End Sub

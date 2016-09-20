'TOC TEMPLATE
Option Explicit
 
 'All of this is to make a table of contents for client reports in the most backwards way possible.
 
Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) _
As Long

'' Win32 API declarations

Public Sub PrintFile(ByVal strPathAndFilename As String)
    Call apiShellExecute(Application.hwnd, "print", strPathAndFilename, vbNullString, vbNullString, 0)
End Sub

Sub TOCFull(list As Variant, tprov As String, tcity As String, ttocpdf As String, ttocxl As String, treportpdf, tprintstring)

    Dim toc_ws As Worksheet
    Dim dump_ws As Worksheet
    Dim ah_ws As Worksheet

    Dim drng As Range
    Dim trng As Range
    Dim hrng As Range
    Dim docname As Range

'I am the woooooorst.

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim m As Integer

    Dim printdocs() As String
    Dim tocdocnames() As String
    Dim tocstring As String
    Dim pstring As String
    Dim filepath As Range

    Set dump_ws = ThisWorkbook.Worksheets("DocHeadings")
    Set toc_ws = ThisWorkbook.Worksheets("TOC")
    Set ah_ws = ThisWorkbook.Worksheets("AllHeadings")

    Set hrng = ah_ws.Range("Table3").Columns(1)

    pstring = tprintstring
    
    i = 0
    For j = LBound(list) To UBound(list)
    'set up dump range to paste each heading from the report
        Set drng = dump_ws.Range("B" & 3 + i)
        drng.Value = Trim(CStr(list(j)))
        'If there is a probate tax heading, add appropriate province to the end.
        If drng.Value = "Probate Tax Rates" Then
            drng.Value = "Probate Tax Rates - " & tprov
        'if there is a heading to consult estate lawyers, add appropriate province to the end.
        ElseIf drng.Value = "Consulting Estate Lawyers" Then
            drng.Value = "Consulting Estate Lawyers - " & tcity
        End If
        For k = 1 To 110
            Set hrng = ah_ws.Range("Table3").Cells(k, 1)
            'once heading strings have been dumped, compare each to pdf file names--if they match, add the file name to the list to be printed
            If drng.Value = hrng.Value Then
                Set docname = hrng.Offset(0, 1)
                Set filepath = hrng.Offset(0, 2)
                'if the list is empty, it's now pstring. If not, add pstring to the end. 
                If filepath.Value <> "" Then
                    If pstring = "" Then
                        pstring = filepath.Value
                    ElseIf pstring <> "" Then
                        pstring = pstring & "," & filepath.Value
                    End If
                    'same thing for the headings to actually include in the finished TOC
                    If tocstring = "" Then
                        tocstring = docname.Value
                    ElseIf tocstring <> "" Then
                        tocstring = tocstring & "," & docname.Value
                    End If
                End If
            End If
        Next k
        If drng.Value <> "" Then
            i = i + 1
        End If
    Next j
    
    'split up the big ol string list into different elements
    printdocs() = Split(pstring, ",")
    tocdocnames() = Split(tocstring, ",")
'    Application.VBE.MainWindow.Visible = True
'set correctly sized range and finally plug pretty headings into it
    For l = LBound(tocdocnames) To UBound(tocdocnames)
        Set trng = toc_ws.Range("B" & 11 + l)
        trng.Value = tocdocnames(l)
    Next l

'save as new worksheet    
    ThisWorkbook.SaveAs Filename:=ttocxl, _
    FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False

'save worsheet as pdf
    toc_ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        ttocpdf, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False

'PRINT ALL THE THINGS (that big list of strings? They're file names. Open those up and print them.)
    For m = LBound(printdocs) To UBound(printdocs)
        PrintFile (printdocs(m))
        'wait 10 seconds--for some reason everything was printing out of order otherwise.
        Application.Wait (Now + TimeValue("0:00:10"))
    Next m

End Sub



''Pretty sure this was to print supplimentary documents without having to creat a new TOC file each time...

''EXTRA
SubPrintSupDocs()
Dim headings_ws as worksheet, toc_ws as worksheet

Dim i as integer, j as integer, k as integer, l as integer
Dim printdocs() as string
Dim headingsrng as range
Dim pstring as string
Dim filepath as range
Dim docnames as range

Set headings_ws = thisworkbook.worksheets("AllHeadings")
Set toc_ws = thisworkbook.worksheets("TOC")
Set headingsrng = headings_ws.range("Table3").Columns(1)

i = toc_ws.Range("B11").End(xlDown).Row
Set docnames = toc_ws.Range("B11:B" & i)
for j = 1 to i - 10
	for k = 1 to 28
		if docnames.cells(j).value = headingsrng.cells(k).value then
			filepath = headingsrng.cells(k).offset(0,2)
			If pstring = "" then
				pstring = filepath.value
			Elseif pstring <> "" then
				pstring = pstring & "," & filepath.value
			End if
		End if
	next k
next j

printdocs() = split(pstring, ",")

    For l = LBound(printdocs) To UBound(printdocs)
        PrintFile (printdocs(l))
        Application.Wait (Now + TimeValue("0:00:10"))
    Next l
End Sub

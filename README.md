# TOCTemplate
TOC template

'TOC TEMPLATE
Option Explicit
 
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
        Set drng = dump_ws.Range("B" & 3 + i)
        drng.Value = Trim(CStr(list(j)))
        If drng.Value = "Probate Tax Rates" Then
            drng.Value = "Probate Tax Rates - " & tprov
        ElseIf drng.Value = "Consulting Estate Lawyers" Then
            drng.Value = "Consulting Estate Lawyers - " & tcity
        End If
        For k = 1 To 110
            Set hrng = ah_ws.Range("Table3").Cells(k, 1)
            If drng.Value = hrng.Value Then
                Set docname = hrng.Offset(0, 1)
                Set filepath = hrng.Offset(0, 2)
                If filepath.Value <> "" Then
                    If pstring = "" Then
                        pstring = filepath.Value
                    ElseIf pstring <> "" Then
                        pstring = pstring & "," & filepath.Value
                    End If
                    If tocstring = "" Then
                        tocstring = docname.Value
                    ElseIf tocstring <> "" Then
                        tocstring = tocstring & "," & docname.Value
                    End If
                End If
            End If
        Next k
        i = i + 1
    Next j
    
    printdocs() = Split(pstring, ",")
    tocdocnames() = Split(tocstring, ",")
'    Application.VBE.MainWindow.Visible = True
    For l = LBound(tocdocnames) To UBound(tocdocnames)
        Set trng = toc_ws.Range("B" & 11 + l)
        trng.Value = tocdocnames(l)
    Next l
    
    ThisWorkbook.SaveAs Filename:=ttocxl, _
    FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
 
    toc_ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        ttocpdf, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False
    
    For m = LBound(printdocs) To UBound(printdocs)
        PrintFile (printdocs(m))
        Application.Wait (Now + TimeValue("0:00:10"))
    Next m
End Sub

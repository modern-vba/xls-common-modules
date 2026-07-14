Attribute VB_Name = "Test_ErrSource"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests that verify class names and member names written in Err.Raise Source.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_ErrRaiseSource_ClassSourceHasModuleAndMemberName(ByVal Assert As UnitTestAssert)
    Dim mismatch_list As ObjectList
    Set mismatch_list = New ObjectList

    Dim vb_proj As Variant
    Set vb_proj = Application.VBE.ActiveVBProject

    Dim vb_comp As Variant
    For Each vb_comp In vb_proj.VBComponents
        If vb_comp.Type = 2 Then
            Call pCollectMismatch(mismatch_list, vb_comp)
        End If
    Next vb_comp

    Assert.Equals "", JoinStringList(mismatch_list)
End Sub

Private Sub pCollectMismatch(ByVal MismatchList As ObjectList, ByVal VbComponent As Variant)
    Dim module_name As String
    module_name = VbComponent.Name

    Dim code_module As Variant
    Set code_module = VbComponent.CodeModule

    Dim line_idx As Long
    For line_idx = 1 To code_module.CountOfLines
        Dim code_line As String
        code_line = code_module.Lines(line_idx, 1)

        Dim source_text As String
        source_text = pExtractClassSourceText(code_line)

        If source_text <> "" Then
            Dim member_sep_pos As Long
            member_sep_pos = InStr(1, source_text, ".", vbBinaryCompare)
            If member_sep_pos = 0 Then
                Call MismatchList.Add(module_name & ":" & line_idx & " Source does not include the member name. (Source=Class " & source_text & ")")
            Else
                Dim source_class_name As String
                source_class_name = Left$(source_text, member_sep_pos - 1)
                If source_class_name <> module_name Then
                    Call MismatchList.Add(module_name & ":" & line_idx & " Source=Class " & source_text)
                End If

                Dim source_member_name As String
                source_member_name = Mid$(source_text, member_sep_pos + 1)
                If source_member_name = "" Then
                    Call MismatchList.Add(module_name & ":" & line_idx & " Source member name is empty. (Source=Class " & source_text & ")")
                End If
            End If
        End If
    Next line_idx
End Sub

Private Function pExtractClassSourceText(ByVal CodeLine As String) As String
    If Left$(LTrim$(CodeLine), 1) = "'" Then Exit Function
    If InStr(1, CodeLine, "Err.Raise", vbBinaryCompare) = 0 Then Exit Function

    Dim marker_text As String
    marker_text = """Class "

    Dim start_pos As Long
    start_pos = InStr(1, CodeLine, marker_text, vbBinaryCompare)
    If start_pos = 0 Then Exit Function

    start_pos = start_pos + Len(marker_text)

    Dim end_pos As Long
    end_pos = InStr(start_pos, CodeLine, """", vbBinaryCompare)
    If end_pos = 0 Then Exit Function

    Dim source_text As String
    source_text = Mid(CodeLine, start_pos, end_pos - start_pos)

    Dim rest_text As String
    rest_text = Mid$(CodeLine, end_pos + 1)
    If Left$(LTrim$(rest_text), 7) = "& ""."" &" Then
        source_text = source_text & ".<dynamic>"
    End If

    pExtractClassSourceText = source_text
End Function

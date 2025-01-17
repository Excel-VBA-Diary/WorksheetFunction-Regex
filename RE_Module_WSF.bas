Attribute VB_Name = "RE_Module_WSF"
Option Explicit

' MIT License
'
' Copyright (c) 2025 Excel-VBA-Diary
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

' Last update: January 17, 2025

'------------------------------------------------------------------------------
' WorksheetFunction REGEXTEST for VBA
' SYNTAX: WSF_REGEXTEST(text, pattern, [case_sensitivity])
'------------------------------------------------------------------------------
Public Function WSF_REGEXTEST(ByVal Text As Variant, _
                              ByVal Pattern As String, _
                              Optional ByVal Case_Sensitivity As Long = 0) As Variant
    Dim result As Variant
    Arg 1, Text
    Arg 2, Pattern
    Arg 3, Case_Sensitivity
    result = [REGEXTEST(Arg(1), Arg(2), Arg(3))]
    If VarType(result) = vbError Then
        WSF_REGEXTEST = ""
    Else
        WSF_REGEXTEST = result
    End If
End Function

'------------------------------------------------------------------------------
' WorksheetFunction REGEXREPLACE for VBA
' SYNTAX: WSF_REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity])
'------------------------------------------------------------------------------
Public Function WSF_REGEXREPLACE(ByVal Text As Variant, _
                                 ByVal Pattern As String, _
                                 ByVal Replacement As String, _
                                 Optional ByVal Occurrence As Long = 0, _
                                 Optional ByVal Case_Sensitivity As Long = 0) As Variant
    Dim result As Variant
    Arg 1, Text
    Arg 2, Pattern
    Arg 3, Replacement
    Arg 4, Occurrence
    Arg 5, Case_Sensitivity
    result = [REGEXREPLACE(Arg(1), Arg(2), Arg(3), Arg(4), Arg(5))]
    If VarType(result) = vbError Then
        WSF_REGEXREPLACE = ""
    Else
        WSF_REGEXREPLACE = result
    End If
End Function

'------------------------------------------------------------------------------
' WorksheetFunction REGEXEXTRACT for VBA
' SYNTAX: WSF_REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])
'------------------------------------------------------------------------------
Public Function WSF_REGEXEXTRACT(ByVal Text As Variant, _
                                 ByVal Pattern As String, _
                                 Optional ByVal Return_Mode As Long = 0, _
                                 Optional ByVal Case_Sensitivity As Long = 0) As Variant
    Dim result As Variant
    Arg 1, Text
    Arg 2, Pattern
    Arg 3, Return_Mode
    Arg 4, Case_Sensitivity
    result = [REGEXEXTRACT(Arg(1), Arg(2), Arg(3), Arg(4))]
    If VarType(result) = vbError Then
        WSF_REGEXEXTRACT = ""
    Else
        WSF_REGEXEXTRACT = result
    End If
End Function

'------------------------------------------------------------------------------
' VBAのデーターをExcel関数の引き数に渡すための関数
' ArgNo  ：書き込みまたは読み出しに使う引き数の番号（1～9まで指定できる）
'          省略した場合は記憶したすべてのデーターを消去する
' ArgData：指定した場合はデーターの書き込み、省略した場合はデーターの読み出し
'------------------------------------------------------------------------------
Private Function Arg(Optional ArgNo As Variant, Optional ArgData As Variant) As Variant
    
    Static temp(1 To 9) As Variant
    
    If IsMissing(ArgNo) Then
        Erase temp
        Exit Function
    End If
    
    If Not IsNumeric(ArgNo) Or ArgNo < 1 Or Arg > 9 Then
        Err.Raise Number:=2001, Description:="Arg: invalid argument"
        Exit Function
    End If
    
    If IsMissing(ArgData) Then
        Arg = temp(ArgNo)               '引き数の読み出し
    Else
        temp(ArgNo) = ArgData           '引き数の書き込み
        Arg = temp(ArgNo)
    End If

End Function

'------------------------------------------------------------------------------
' End of Source Code
'------------------------------------------------------------------------------


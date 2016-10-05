Attribute VB_Name = "excelFunction"
' join �w�肵���͈͂̕�������������܂�
' @param sel �͈�
' @param delimiter �������̋�؂蕶���B�w��Ȃ��̏ꍇ�A�󕶎�
' @param skipBlank ��̃Z�����X�L�b�v���邩�B�w��Ȃ��̏ꍇ�A�X�L�b�v����
' @param skipHiddenCell ��\���̃Z�����X�L�b�v���邩�B�w��Ȃ��̏ꍇ�A�X�L�b�v����
Function join(sel As Range, Optional delimiter As String = "", Optional skipBlank As Boolean = True, Optional skipHiddenCell As Boolean = True) As String
    Dim rng As Range
    Dim ret As String
    Dim initFlag As Boolean
    Dim delim As String
    
    initFlag = True
    ret = ""

    delim = delimiter
    
    For Each rng In sel
        If rng.Text = "" And skipBlank = True Then
        ElseIf skipHiddenCell = True And (rng.Height = 0 Or rng.Width = 0) Then
        Else
            If initFlag = False Then
                ret = ret & delim
            Else
                initFlag = False
            End If
        
            ret = ret & rng.Text
        End If
    Next
    
    join = ret

End Function

' decode �I���N����decode�Ɠ���
Function decode(sel As Variant, k1 As String, v1 As String, ParamArray varArray() As Variant)
    Dim dct As Object
    Dim i As Integer
    Dim key As String
    Dim keystr As String
    
    Select Case TypeName(sel)
    Case "String"
        keystr = sel
    Case "Range"
        keystr = sel.Value
    Case Else
        keystr = ""
    End Select
    
    If sel.Value = k1 Then
        decode = v1
        GoTo endf
    End If
    
    Set dct = CreateObject("Scripting.Dictionary")
    For i = LBound(varArray) To UBound(varArray)
        
        If i Mod 2 = 0 Then
            key = varArray(i)
        Else
            dct.Add key, varArray(i)
            key = ""
        End If
    Next i
    
    If dct.Exists(sel.Value) Then
        decode = dct.Item(sel.Value)
    Else
        decode = key
    End If
    
    
endf:
End Function

' sp ������𕪉����A�C���f�b�N�X�Ŏw�肵���Ԗڂ̒l��Ԃ�
' @param r �����񂩕����񂪊i�[���ꂽ�Z��
' @param idx 0����n�܂�C���f�b�N�X
' @param delim ��؂蕶��(��)�B�w��Ȃ��̏ꍇ�A" "
Function sp(r As Variant, idx As Integer, Optional delim As String = " ")
    Dim ptn As String
    Dim re As Object
    Dim str As String
    
    Set re = CreateObject("VBScript.RegExp")
    
    Select Case TypeName(r)
    Case "String"
        str = r
    Case "Range"
        str = r.Value
    Case Else
        str = ""
    End Select
    
    ptn = Replace(delim, "\", "\\")
    ptn = Replace(ptn, "+", "\+")
    ptn = Replace(ptn, ".", "\.")
    ptn = Replace(ptn, "(", "\(")
    ptn = Replace(ptn, ")", "\)")
    ptn = Replace(ptn, "{", "\{")
    ptn = Replace(ptn, "}", "\}")
    ptn = Replace(ptn, "[", "\[")
    ptn = Replace(ptn, "-", "\-")
    ptn = Replace(ptn, "^", "\^")
    ptn = Replace(ptn, "*", "\*")
    
    
    re.Pattern = "(" & ptn & ")+"
    re.Global = True
    
    sp = Trim(split(re.Replace(str, "_"), "_")(idx))
End Function

Attribute VB_Name = "excelFunction"
' join 指定した範囲の文字列を結合します
' @param sel 範囲
' @param delimiter 結合時の区切り文字。指定なしの場合、空文字
' @param skipBlank 空のセルをスキップするか。指定なしの場合、スキップする
' @param skipHiddenCell 非表示のセルをスキップするか。指定なしの場合、スキップする
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

' decode オラクルのdecodeと同じ
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

' sp 文字列を分解し、インデックスで指定した番目の値を返す
' @param r 文字列か文字列が格納されたセル
' @param idx 0から始まるインデックス
' @param delim 区切り文字(列)。指定なしの場合、" "
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

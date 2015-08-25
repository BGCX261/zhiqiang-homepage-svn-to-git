Attribute VB_Name = "GenFunList"
Option Explicit

Sub GenerateFunList()
    Dim fl$
    fl = "D:\My Documents\07.����ģ��\excel\CData.cls"
    
    Open fl For Input As #1
    Dim textLine$, oneItem
    
    Dim className$
    Do
        Line Input #1, textLine
    Loop Until EOF(1) Or VBA.Mid(textLine, 1, 1) = "'"
    className = VBA.Trim(VBA.Mid(textLine, 2, 1000))
    
    Dim funName$, funExp$, funClass$, funDef$, funPara$(), regCode$
    regCode = "Public Sub RegFun()"
    Do While Not EOF(1)
        ' ���뵽һ������Ϊֹ�������ɺ�����ע��
        Line Input #1, textLine
        Do While InStr(1, textLine, "Public Function") = 0 And InStr(1, textLine, "Public sub") = 0 And Not EOF(1)
            textLine = VBA.Trim(textLine)
            If VBA.Mid(textLine, 1, 1) = "'" Then
                funExp = funExp & IIf(Len(funExp), vbCrLf, "") & textLine
            Else
                funExp = ""
            End If
            Line Input #1, textLine
        Loop
        ' If Len(funExp) Then funExp = VBA.Mid(funExp, 2)
        
        ' ����Ѿ��Ҳ����º��������Ƴ�ѭ��
        If EOF(1) Then Exit Do
        
        ' ȷ����������
        If InStr(1, textLine, "Public Function") = 1 Then
            funClass = "Function"
        ElseIf InStr(1, textLine, "Public Sub") = 1 Then
            funClass = "Sub"
        Else
            Exit Do
        End If
        
        ' ���뺯������
        funDef = textLine
        ' ��Ҫ�����л��з�������
        Do While VBA.Mid(textLine, Len(textLine) - 1, 2) = " _" And Not EOF(1)
            ' ɾ�����з�
            funDef = VBA.Mid(funDef, 1, Len(funDef) - 2)
            
            ' ����������һ��
            Line Input #1, textLine
            textLine = VBA.Trim(textLine)
            funDef = funDef & textLine
            
            If Len(textLine) < 2 Then Exit Do
        Loop
        Dim cleanDef$, paraArray$()
        ' ɾ�������е��ַ�������Ҫ�ǿ�ѡ������Ĭ��ֵ�������ַ���
        Dim i&, ind As Boolean
        ind = False
        cleanDef = ""
        For i = 1 To Len(funDef)
            If ind = False Then
                ' ��Ч�ַ�
                If VBA.Mid(funDef, i, 1) <> """" Then
                    cleanDef = cleanDef & VBA.Mid(funDef, i, 1)
                Else
                ' �����ַ���ģʽ
                    ind = True
                End If
            Else
                ' �˳��ַ���ģʽ
                If VBA.Mid(funDef, i, 1) = """" Then
                    ind = False
                End If
            End If
        Next i
        
        ' ��ȡ�����ַ������飬��󱣴���funPara������
        Dim paraStr$
        paraStr = VBA.Mid(funDef, InStr(1, funDef, "(") + 1, _
                InStr(1, funDef, ")") - InStr(1, funDef, "(") - 1)
        Dim paraStrArr$(), tmpStr, tmpArr$()
        paraStrArr = Split(paraStr, ",")
        ReDim funPara(0 To UBound(paraStrArr))
        For i = 0 To UBound(paraStrArr)
            tmpStr = VBA.Trim(paraStrArr(i))
            tmpArr = Split(tmpStr, " ")
            For Each tmpStr In tmpArr
                If Not tmpStr = "Optional" And Not tmpStr = "ByVal" And Not tmpStr = "ByRef" Then
                    funPara(i) = tmpStr
                    Dim cs As String
                    cs = VBA.Right(tmpStr, 1)
                    If cs = "&" Or cs = "$" Then
                        funPara(i) = VBA.Left(tmpStr, Len(tmpStr) - 1)
                    End If
                    
                    Exit For
                End If
            Next tmpStr
            
        Next i
        
        ' ��ȡ������
        tmpStr = VBA.Trim(cleanDef)
        tmpStr = VBA.Left(tmpStr, InStr(1, tmpStr, "(") - 1)
        tmpArr = Split(tmpStr, " ")
        For Each tmpStr In tmpArr
            If Not tmpStr = "Sub" And Not tmpStr = "Function" And Not tmpStr = "Public" Then
                funName = tmpStr
                cs = VBA.Right(tmpStr, 1)
                If cs = "&" Or cs = "$" Then
                    funName = VBA.Left(tmpStr, Len(tmpStr) - 1)
                End If
                
                Exit For
            End If
        Next tmpStr
        
        ' �ҵ�һ�������󣬼������뺯��bodyֱ����������
        Do While Not EOF(1) And InStr(1, textLine, "End Sub") <> 1 And InStr(1, textLine, "End Function") <> 1
            Line Input #1, textLine
            textLine = VBA.Trim(textLine)
        Loop
        
        Debug.Print vbCrLf & funExp
        Debug.Print funDef
        Debug.Print "    " & IIf(funClass = "Sub", "Call ", funName & " = ") & className & "." & funName & "(" & Join(funPara, ", ") & ")"
        Debug.Print "    " & "If(conDegug) then debug.print now() & " & """" & funName & """ & ""("" & " & Join(funPara, " & "", "" & ") & """)"""
        Debug.Print "End " & funClass
        
        regCode = regCode & vbCrLf & "    " & "Call Application.MacroOptions(""" & StringForCode((funName)) & """, """ & _
                 StringForCode(funExp) & """, , , , , , ""����֤ȯ"")"
    Loop
    
    regCode = regCode & vbCrLf & "End Sub"
    Debug.Print regCode
    Close #1
End Sub


' ���ַ���ת����vb������õ��ַ���
Private Function StringForCode$(strIn$)
    StringForCode = Replace(strIn, """", """""")
    StringForCode = Replace(StringForCode, vbCrLf, " "" & vbCrLf & """)
End Function

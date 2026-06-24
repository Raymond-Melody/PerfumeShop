<%
' ============================================
' V14.6 公共工具函数库
' 提供全局通用的安全类型转换和辅助函数
' 引入方式: <!--#include file="../../includes/common_utils.asp"-->
' ============================================

' ========== SafeNum：安全数值转换 ==========
' 将任意值安全转换为双精度浮点数，无效值返回0
Function SafeNum(val)
    On Error Resume Next
    If IsNull(val) Or IsEmpty(val) Or val = "" Then
        SafeNum = 0
    ElseIf Not IsNumeric(val) Then
        SafeNum = 0
    Else
        SafeNum = CDbl(val)
        If Err.Number <> 0 Then
            SafeNum = 0
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Function

' ========== SafeInt：安全整数转换 ==========
' 将任意值安全转换为整数，无效值返回0
Function SafeInt(val)
    On Error Resume Next
    If IsNull(val) Or IsEmpty(val) Or val = "" Then
        SafeInt = 0
    ElseIf Not IsNumeric(val) Then
        SafeInt = 0
    Else
        SafeInt = CLng(val)
        If Err.Number <> 0 Then
            SafeInt = 0
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Function

' ========== SafeFloat：安全浮点转换 ==========
' 与SafeNum相同但语义更明确，用于需要浮点数的场景
Function SafeFloat(val)
    SafeFloat = SafeNum(val)
End Function

' ========== SafeDiv：安全除法 ==========
' 防止除零错误，分母无效时返回0
Function SafeDiv(numerator, denominator)
    On Error Resume Next
    If IsNull(denominator) Or denominator = "" Then
        SafeDiv = 0
    ElseIf Not IsNumeric(denominator) Then
        SafeDiv = 0
    ElseIf CDbl(denominator) = 0 Then
        SafeDiv = 0
    Else
        SafeDiv = CDbl(numerator) / CDbl(denominator)
        If Err.Number <> 0 Then
            SafeDiv = 0
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Function

' ========== IIF：条件表达式 ==========
' 类似三元运算符，cond为True返回tVal，否则返回fVal
Function IIF(cond, tVal, fVal)
    If cond Then IIF = tVal Else IIF = fVal
End Function

' ========== SafeStr：安全字符串转换 ==========
' 将任意值安全转换为字符串，Null/Empty返回空字符串
Function SafeStr(val)
    If IsNull(val) Or IsEmpty(val) Then
        SafeStr = ""
    Else
        SafeStr = CStr(val)
    End If
End Function

' ========== SafeBool：安全布尔转换 ==========
' 将任意值安全转换为布尔值
Function SafeBool(val)
    On Error Resume Next
    If IsNull(val) Or IsEmpty(val) Or val = "" Then
        SafeBool = False
    ElseIf IsNumeric(val) Then
        SafeBool = (CDbl(val) <> 0)
    ElseIf LCase(CStr(val)) = "true" Or LCase(CStr(val)) = "yes" Then
        SafeBool = True
    Else
        SafeBool = False
    End If
    On Error GoTo 0
End Function

' ========== FormatPrice：格式化价格 ==========
' 格式化金额为指定小数位数
Function FormatPrice(val, decimals)
    On Error Resume Next
    If Not IsNumeric(decimals) Then decimals = 2
    FormatPrice = FormatNumber(SafeNum(val), decimals)
    If Err.Number <> 0 Then
        FormatPrice = "0.00"
        Err.Clear
    End If
    On Error GoTo 0
End Function
%>

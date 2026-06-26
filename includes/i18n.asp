<%
' ============================================
' V15.0 国际化基础 (i18n)
' 依赖: 无
' 用法: <!--#include file="i18n.asp"-->
' 调用: Response.Write T("welcome_message")
'        Response.Write T("order_count", Array(count))  ' 带参数
' 语言检测优先级: URL ?lang= → Session → Cookie → 浏览器Accept-Language → 默认zh-CN
' ============================================

' 支持的语言
Const I18N_DEFAULT_LOCALE = "zh-CN"
Const I18N_SUPPORTED_LOCALES = "zh-CN,en-US"

' 语言字典（Application缓存）
Dim I18N_CurrentLocale

' ============================================
' 检测当前语言
' ============================================
Function I18N_DetectLocale()
    Dim locale, acceptLang
    
    ' 1. URL参数优先
    locale = Request.QueryString("lang")
    If locale <> "" Then
        Session("Locale") = locale
        I18N_DetectLocale = locale
        Exit Function
    End If
    
    ' 2. Session缓存
    locale = Session("Locale")
    If locale <> "" Then
        I18N_DetectLocale = locale
        Exit Function
    End If
    
    ' 3. Cookie
    locale = Request.Cookies("lang")
    If locale <> "" Then
        I18N_DetectLocale = locale
        Exit Function
    End If
    
    ' 4. 浏览器 Accept-Language
    acceptLang = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
    If acceptLang <> "" Then
        ' 检查是否包含中文
        If InStr(1, acceptLang, "zh", vbTextCompare) > 0 Then
            I18N_DetectLocale = "zh-CN"
            Exit Function
        End If
        ' 检查是否包含英文
        If InStr(1, acceptLang, "en", vbTextCompare) > 0 Then
            I18N_DetectLocale = "en-US"
            Exit Function
        End If
    End If
    
    ' 5. 默认
    I18N_DetectLocale = I18N_DEFAULT_LOCALE
End Function

' ============================================
' 加载语言包到内存
' ============================================
Function I18N_LoadDictionary(locale)
    Dim cacheKey, dict, fso, filePath, line, pos
    
    cacheKey = "I18N_DICT_" & locale
    
    ' 先从Application缓存获取
    If IsObject(Application(cacheKey)) Then
        Set I18N_LoadDictionary = Application(cacheKey)
        Exit Function
    End If
    
    ' 从文件加载（使用 ADODB.Stream 读取 UTF-8 文件）
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    filePath = Server.MapPath("/locale/" & locale & ".asp")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Err.Number = 0 And fso.FileExists(filePath) Then
        ' V17.2: 使用 ADODB.Stream 以 UTF-8 编码读取 locale 文件
        ' FSO.OpenTextFile 在中文系统默认以 GBK 读取，导致 UTF-8 中文丢失
        Dim stream, content, lines, lineIdx
        Set stream = Server.CreateObject("ADODB.Stream")
        stream.Type = 2  ' adTypeText
        stream.Charset = "UTF-8"
        stream.Open
        stream.LoadFromFile filePath
        content = stream.ReadText
        stream.Close
        Set stream = Nothing
        
        ' 按行分割（兼容 CRLF 和 LF）
        lines = Split(content, vbCrLf)
        If UBound(lines) < 0 Then lines = Split(content, vbLf)
        For lineIdx = 0 To UBound(lines)
            line = Trim(lines(lineIdx))
            ' 格式: key=value
            If line <> "" And Left(line, 1) <> "#" Then
                pos = InStr(line, "=")
                If pos > 0 Then
                    dict.Add Trim(Left(line, pos - 1)), Trim(Mid(line, pos + 1))
                End If
            End If
        Next
    End If
    
    Set fso = Nothing
    Err.Clear
    
    ' 缓存到Application
    Application.Lock
    If Not IsObject(Application(cacheKey)) Then
        Set Application(cacheKey) = dict
    End If
    Application.UnLock
    
    Set I18N_LoadDictionary = dict
End Function

' ============================================
' 翻译函数 T(key, params)
' ============================================
Function T(key, params)
    Dim dict, value
    
    ' 使用当前语言
    If IsEmpty(I18N_CurrentLocale) Or I18N_CurrentLocale = "" Then
        I18N_CurrentLocale = I18N_DetectLocale()
    End If
    
    Set dict = I18N_LoadDictionary(I18N_CurrentLocale)
    
    If dict.Exists(key) Then
        value = dict.Item(key)
    ElseIf I18N_CurrentLocale <> I18N_DEFAULT_LOCALE Then
        ' 回退到默认语言
        Dim defaultDict
        Set defaultDict = I18N_LoadDictionary(I18N_DEFAULT_LOCALE)
        If defaultDict.Exists(key) Then
            value = defaultDict.Item(key)
        Else
            T = "[" & key & "]"
            Exit Function
        End If
        Set defaultDict = Nothing
    Else
        T = "[" & key & "]"
        Exit Function
    End If
    
    ' 替换参数 {0}, {1}, ...
    If IsArray(params) Then
        Dim i
        For i = 0 To UBound(params)
            value = Replace(value, "{" & i & "}", CStr(params(i)))
        Next
    ElseIf Not IsNull(params) And Not IsEmpty(params) Then
        value = Replace(value, "{0}", CStr(params))
    End If
    
    T = value
End Function

' ============================================
' 设置当前语言
' ============================================
Sub I18N_SetLocale(locale)
    I18N_CurrentLocale = locale
    Session("Locale") = locale
    Response.Cookies("lang") = locale
    Response.Cookies("lang").Expires = DateAdd("yyyy", 1, Now())
End Sub

' ============================================
' 获取当前语言
' ============================================
Function I18N_GetLocale()
    If IsEmpty(I18N_CurrentLocale) Or I18N_CurrentLocale = "" Then
        I18N_CurrentLocale = I18N_DetectLocale()
    End If
    I18N_GetLocale = I18N_CurrentLocale
End Function

' ============================================
' 获取HTML lang属性值
' ============================================
Function I18N_HtmlLang()
    I18N_HtmlLang = I18N_GetLocale()
End Function

' ============================================
' 初始化
' ============================================
I18N_CurrentLocale = I18N_DetectLocale()
%>
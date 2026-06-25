<%
' ============================================
' V14.6 结算页 - 工具函数
' 从 checkout.asp 提取
' ============================================

' 调试日志函数
Sub DebugLog(msg)
    Dim fso, logFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.OpenTextFile(Server.MapPath("/debug_checkout.log"), 8, True)
    logFile.WriteLine Now() & " - " & msg
    logFile.Close
    Set logFile = Nothing
    Set fso = Nothing
End Sub

' 通用成分分割函数 - 支持逗号、空格、换行符等多种分隔符
Function SplitIngredientsUniversal(rawStr)
    Dim result, arr, item, i
    Set result = CreateObject("Scripting.Dictionary")

    If rawStr = "" Then
        Set SplitIngredientsUniversal = result
        Exit Function
    End If

    ' 统一将所有分隔符转换为英文逗号
    rawStr = Replace(rawStr, "，", ",")      ' 中文逗号
    rawStr = Replace(rawStr, vbCrLf, ",")   ' 回车换行
    rawStr = Replace(rawStr, vbLf, ",")     ' 换行符
    rawStr = Replace(rawStr, vbCr, ",")     ' 回车符
    rawStr = Replace(rawStr, Chr(160), ",") ' NBSP
    rawStr = Replace(rawStr, " ", ",")     ' 全角空格

    ' 清理连续逗号
    Do While InStr(rawStr, ",,") > 0
        rawStr = Replace(rawStr, ",,", ",")
    Loop

    ' 用逗号分割
    arr = Split(rawStr, ",")
    For i = 0 To UBound(arr)
        item = Trim(arr(i))
        If item <> "" And Not result.Exists(item) Then
            result.Add item, True
        End If
    Next

    Set SplitIngredientsUniversal = result
End Function

' 构建完整地址（自动去重城市/区域）
Function BuildFullAddress(province, city, district, detail)
    Dim result
    province = Trim(province & "")
    city = Trim(city & "")
    district = Trim(district & "")
    detail = Trim(detail & "")

    result = province
    If city <> "" Then
        ' 去重：如果city和district相同，只保留city
        If city = district Then
            result = result & city
        Else
            result = result & city
            If district <> "" Then
                result = result & district
            End If
        End If
    End If
    If detail <> "" Then
        result = result & detail
    End If
    BuildFullAddress = result
End Function

' V17: 根据ID获取地区名称的函数 - 使用参数化查询
Function GetAreaNameById(areaId)
    If Not IsNumeric(areaId) Or areaId = "" Then
        GetAreaNameById = ""
        Exit Function
    End If

    Dim sql, rs, params(0)
    sql = "SELECT AreaName FROM Areas WHERE AreaID = @AreaID"
    params(0) = Array("@AreaID", DAL_adInteger, 0, CLng(areaId))
    Set rs = DAL_GetList(sql, params)

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            GetAreaNameById = rs("AreaName")
        Else
            GetAreaNameById = ""
        End If
        rs.Close
        Set rs = Nothing
    Else
        GetAreaNameById = ""
    End If
End Function
%>

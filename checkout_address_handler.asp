<%
' ============================================
' V14.6 结算页 - 地址POST处理器
' 从 checkout.asp 提取
' ============================================

' 确保CSRF令牌存在
Call EnsureCSRFToken()

' CSRF验证
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        Response.Write "<script>alert('安全验证失败，请刷新页面重试'); history.back();</script>"
        Response.End
    End If
End If

' 处理表单提交 - 添加地址
If Request.Form("action") = "add" Then
    Dim consignee, phoneNum, provinceName, cityName, districtName, detailAddress, isDefaultAddr
    consignee = SafeSQL(Trim(Request.Form("realName")))
    phoneNum = SafeSQL(Trim(Request.Form("phone")))
    provinceName = SafeSQL(Trim(Request.Form("province")))
    cityName = SafeSQL(Trim(Request.Form("city")))
    districtName = SafeSQL(Trim(Request.Form("district")))
    detailAddress = SafeSQL(Trim(Request.Form("address")))
    isDefaultAddr = Request.Form("isDefault")

    If isDefaultAddr <> "" And isDefaultAddr <> "0" Then
        isDefaultAddr = 1
    Else
        isDefaultAddr = 0
    End If

    ' 验证收货信息
    If consignee = "" Or phoneNum = "" Or provinceName = "" Or cityName = "" Or districtName = "" Or detailAddress = "" Then
        Session("ErrorMessage") = "请填写完整的收货信息"
    Else
        ' 如果设为默认地址，先取消其他默认地址
        If isDefaultAddr <> 0 Then
            Call ExecuteNonQuery("UPDATE UserAddresses SET IsDefault = 0 WHERE UserID = " & userId)
        End If

        Dim insertSql
        insertSql = "INSERT INTO UserAddresses (UserID, Consignee, Phone, Province, City, District, Address, IsDefault, CreatedAt) VALUES (" & userId & ", '" & consignee & "', '" & phoneNum & "', '" & provinceName & "', '" & cityName & "', '" & districtName & "', '" & detailAddress & "', " & isDefaultAddr & ", GETDATE())"

        If ExecuteNonQuery(insertSql) Then
            Dim newAddressId
            newAddressId = GetLastInsertID("UserAddresses")
            Dim paymentMethodFromForm
            paymentMethodFromForm = Request.Form("payment_method")
            Dim redirectUrl
            redirectUrl = "checkout.asp"

            Dim queryString
            queryString = ""

            If cartIds <> "" Then
                queryString = queryString & "cart_ids=" & cartIds
            End If

            If paymentMethodFromForm <> "" Then
                If queryString <> "" Then queryString = queryString & "&"
                queryString = queryString & "payment_method=" & paymentMethodFromForm
            End If

            If isDefaultAddr <> 0 Then
                If queryString <> "" Then queryString = queryString & "&"
                queryString = queryString & "selected_address=" & newAddressId
            End If

            If queryString <> "" Then
                redirectUrl = redirectUrl & "?" & queryString
            End If

            Response.Redirect redirectUrl
            Response.End
        Else
            Session("ErrorMessage") = "地址保存失败，请重试"
        End If
    End If
End If
%>

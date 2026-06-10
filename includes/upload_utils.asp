<%
' ============================================
' Multipart/Form-Data 解析与文件上传工具库
' 纯 ASP Classic 实现，无外部依赖
' 被其他 ASP 页面 include 使用，不应有任何 HTML 输出
' ============================================

' --------------------------------------------
' 常量定义
' --------------------------------------------
Const MAX_UPLOAD_SIZE = 5242880
Const ALLOWED_IMAGE_EXTENSIONS = ".jpg,.jpeg,.png,.gif,.webp,.svg"

' --------------------------------------------
' MultipartParser 类
' --------------------------------------------
Class MultipartParser
    Private m_formFields
    Private m_fileName
    Private m_fileData
    Private m_fileSize
    Private m_contentType
    Private m_boundary
    Private m_errorMsg

    Private Sub Class_Initialize()
        Set m_formFields = Server.CreateObject("Scripting.Dictionary")
        m_fileName = ""
        Set m_fileData = Nothing
        m_fileSize = 0
        m_contentType = ""
        m_boundary = ""
        m_errorMsg = ""
    End Sub

    Private Sub Class_Terminate()
        Set m_formFields = Nothing
        If Not m_fileData Is Nothing Then
            m_fileData.Close
            Set m_fileData = Nothing
        End If
    End Sub

    ' 解析 multipart/form-data 请求
    Public Sub Parse()
        Dim contentType, totalBytes
        Dim reqStream, txtStream, textData
        Dim bMark, pos, nextPos
        Dim partStart, partLen, partText
        Dim headerEnd, headerText
        Dim bodyStart, bodyLen
        Dim fieldName, fileName, fContentType
        Dim fieldBytes, fieldStream, fieldValue

        On Error Resume Next

        contentType = Request.ServerVariables("CONTENT_TYPE")
        If Len(contentType) = 0 Then
            m_errorMsg = "无法获取 Content-Type"
            Exit Sub
        End If

        pos = InStr(LCase(contentType), "boundary=")
        If pos = 0 Then
            m_errorMsg = "Content-Type 中未找到 boundary"
            Exit Sub
        End If

        m_boundary = Trim(Mid(contentType, pos + 9))
        If Left(m_boundary, 1) = """" And Right(m_boundary, 1) = """" Then
            m_boundary = Mid(m_boundary, 2, Len(m_boundary) - 2)
        End If

        If Len(m_boundary) = 0 Then
            m_errorMsg = "boundary 为空"
            Exit Sub
        End If

        totalBytes = Request.TotalBytes
        If totalBytes = 0 Then
            m_errorMsg = "请求体为空"
            Exit Sub
        End If

        If totalBytes > MAX_UPLOAD_SIZE Then
            m_errorMsg = "上传文件超过大小限制 " & MAX_UPLOAD_SIZE & " 字节"
            Exit Sub
        End If

        Set reqStream = Server.CreateObject("ADODB.Stream")
        reqStream.Type = 1
        reqStream.Open
        reqStream.Write Request.BinaryRead(totalBytes)

        If Err.Number <> 0 Then
            If InStr(Err.Description, "0104") > 0 Or InStr(LCase(Err.Description), "不允许") > 0 Then
                m_errorMsg = "上传文件过大或请求体读取被阻止。请检查文件大小是否超过服务器限制（最大 " & (MAX_UPLOAD_SIZE \ 1048576) & "MB）。错误: ASP 0104"
            Else
                m_errorMsg = "读取请求体失败: " & Err.Description
            End If
            reqStream.Close
            Set reqStream = Nothing
            Exit Sub
        End If

        reqStream.Position = 0
        Set txtStream = Server.CreateObject("ADODB.Stream")
        txtStream.Type = 1
        txtStream.Open
        reqStream.CopyTo txtStream
        txtStream.Position = 0
        txtStream.Type = 2
        txtStream.Charset = "iso-8859-1"
        textData = txtStream.ReadText
        txtStream.Close
        Set txtStream = Nothing

        bMark = "--" & m_boundary
        pos = 1

        Do While pos > 0
            pos = InStr(pos, textData, bMark)
            If pos = 0 Then Exit Do

            If Mid(textData, pos + Len(bMark), 2) = "--" Then Exit Do

            nextPos = InStr(pos + Len(bMark), textData, bMark)

            partStart = pos + Len(bMark) + Len(vbCrLf)
            If nextPos > 0 Then
                partLen = nextPos - partStart - Len(vbCrLf)
            Else
                partLen = Len(textData) - partStart + 1
            End If

            If partLen > 0 Then
                partText = Mid(textData, partStart, partLen)

                headerEnd = InStr(partText, vbCrLf & vbCrLf)
                If headerEnd > 0 Then
                    headerText = Left(partText, headerEnd - 1)

                    bodyStart = partStart + headerEnd + Len(vbCrLf & vbCrLf) - 1
                    bodyLen = partLen - headerEnd - Len(vbCrLf & vbCrLf) + 1

                    If bodyLen > 0 Then
                        Call ParseDispositionHeader(headerText, fieldName, fileName, fContentType)

                        If Len(fileName) > 0 Then
                            m_fileName = fileName
                            m_contentType = fContentType
                            m_fileSize = bodyLen

                            Set m_fileData = Server.CreateObject("ADODB.Stream")
                            m_fileData.Type = 1
                            m_fileData.Open
                            reqStream.Position = bodyStart - 1
                            m_fileData.Write reqStream.Read(bodyLen)
                        ElseIf Len(fieldName) > 0 Then
                            reqStream.Position = bodyStart - 1
                            fieldBytes = reqStream.Read(bodyLen)

                            Set fieldStream = Server.CreateObject("ADODB.Stream")
                            fieldStream.Type = 1
                            fieldStream.Open
                            fieldStream.Write fieldBytes
                            fieldStream.Position = 0
                            fieldStream.Type = 2
                            fieldStream.Charset = "UTF-8"
                            fieldValue = fieldStream.ReadText
                            fieldStream.Close
                            Set fieldStream = Nothing

                            m_formFields(fieldName) = fieldValue
                        End If
                    End If
                End If
            End If

            pos = nextPos
        Loop

        reqStream.Close
        Set reqStream = Nothing

        On Error GoTo 0
    End Sub

    ' 解析 Content-Disposition header
    Private Sub ParseDispositionHeader(headerText, ByRef fieldName, ByRef fileName, ByRef fContentType)
        Dim lines, i, line
        Dim pos1, pos2

        fieldName = ""
        fileName = ""
        fContentType = ""

        lines = Split(headerText, vbCrLf)
        For i = 0 To UBound(lines)
            line = lines(i)

            If InStr(LCase(line), "content-disposition:") > 0 Then
                pos1 = InStr(LCase(line), "name=")
                If pos1 > 0 Then
                    pos1 = pos1 + 5
                    If Mid(line, pos1, 1) = """" Then
                        pos1 = pos1 + 1
                        pos2 = InStr(pos1, line, """")
                        If pos2 > 0 Then
                            fieldName = Mid(line, pos1, pos2 - pos1)
                        End If
                    Else
                        pos2 = InStr(pos1, line, ";")
                        If pos2 = 0 Then pos2 = Len(line) + 1
                        fieldName = Trim(Mid(line, pos1, pos2 - pos1))
                    End If
                End If

                pos1 = InStr(LCase(line), "filename=")
                If pos1 > 0 Then
                    pos1 = pos1 + 9
                    If Mid(line, pos1, 1) = """" Then
                        pos1 = pos1 + 1
                        pos2 = InStr(pos1, line, """")
                        If pos2 > 0 Then
                            fileName = Mid(line, pos1, pos2 - pos1)
                        End If
                    Else
                        pos2 = InStr(pos1, line, ";")
                        If pos2 = 0 Then pos2 = Len(line) + 1
                        fileName = Trim(Mid(line, pos1, pos2 - pos1))
                    End If

                    If InStrRev(fileName, "\") > 0 Then
                        fileName = Mid(fileName, InStrRev(fileName, "\") + 1)
                    End If
                    If InStrRev(fileName, "/") > 0 Then
                        fileName = Mid(fileName, InStrRev(fileName, "/") + 1)
                    End If
                End If
            ElseIf InStr(LCase(line), "content-type:") > 0 Then
                pos1 = InStr(line, ":")
                If pos1 > 0 Then
                    fContentType = Trim(Mid(line, pos1 + 1))
                End If
            End If
        Next
    End Sub

    Public Property Get FormField(name)
        If m_formFields.Exists(name) Then
            FormField = m_formFields(name)
        Else
            FormField = ""
        End If
    End Property

    Public Property Get FileName()
        FileName = m_fileName
    End Property

    Public Property Get FileData()
        Set FileData = m_fileData
    End Property

    Public Property Get FileSize()
        FileSize = m_fileSize
    End Property

    Public Property Get ContentType()
        ContentType = m_contentType
    End Property

    Public Property Get HasFile()
        HasFile = (Len(m_fileName) > 0 And Not m_fileData Is Nothing)
    End Property

    Public Property Get ErrorMsg()
        ErrorMsg = m_errorMsg
    End Property
End Class

' --------------------------------------------
' 保存上传的文件
' --------------------------------------------
Function SaveUploadedFile(binaryData, filePath)
    Dim stream

    On Error Resume Next
    SaveUploadedFile = False

    If binaryData Is Nothing Then
        Exit Function
    End If

    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open

    binaryData.Position = 0
    stream.Write binaryData.Read()
    stream.SaveToFile filePath, 2

    If Err.Number = 0 Then
        SaveUploadedFile = True
    End If

    stream.Close
    Set stream = Nothing
    On Error GoTo 0
End Function

' --------------------------------------------
' 生成唯一文件名
' --------------------------------------------
Function GenerateUploadFileName(originalName)
    Dim ext, fileName, randomNum

    ext = ""
    If InStrRev(originalName, ".") > 0 Then
        ext = Mid(originalName, InStrRev(originalName, "."))
    End If

    Randomize
    randomNum = Int(90000 * Rnd) + 10000

    fileName = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & _
               Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & _
               "_" & randomNum & ext

    GenerateUploadFileName = fileName
End Function

' --------------------------------------------
' 验证是否为允许的图片类型
' --------------------------------------------
Function IsValidImageType(fileName)
    Dim ext

    ext = LCase(fileName)
    If InStrRev(ext, ".") > 0 Then
        ext = Mid(ext, InStrRev(ext, "."))
    Else
        IsValidImageType = False
        Exit Function
    End If

    IsValidImageType = (InStr(ALLOWED_IMAGE_EXTENSIONS, ext) > 0)
End Function

' --------------------------------------------
' 验证文件头魔数
' --------------------------------------------
Function IsValidImageMagicBytes(binaryData)
    Dim bytes, allBytes
    Dim isJPEG, isPNG, isGIF, isWebP, isBMP

    On Error Resume Next
    IsValidImageMagicBytes = False

    If binaryData Is Nothing Then Exit Function

    binaryData.Position = 0
    ' 读取12字节以支持WebP检测（RIFF+size+WEBP）
    allBytes = binaryData.Read(12)
    binaryData.Position = 0

    If Err.Number <> 0 Then Exit Function

    If UBound(allBytes) < 2 Then Exit Function

    ' 复制前8字节到bytes以兼容原有检查
    ReDim bytes(7)
    Dim i
    For i = 0 To 7
        If i <= UBound(allBytes) Then
            bytes(i) = allBytes(i)
        End If
    Next

    isJPEG = (bytes(0) = &HFF And bytes(1) = &HD8 And bytes(2) = &HFF)

    isPNG = False
    If UBound(bytes) >= 3 Then
        isPNG = (bytes(0) = &H89 And bytes(1) = &H50 And bytes(2) = &H4E And bytes(3) = &H47)
    End If

    isGIF = False
    If UBound(bytes) >= 3 Then
        isGIF = (bytes(0) = &H47 And bytes(1) = &H49 And bytes(2) = &H46 And bytes(3) = &H38)
    End If

    isBMP = False
    If UBound(bytes) >= 1 Then
        isBMP = (bytes(0) = &H42 And bytes(1) = &H4D)
    End If

    ' WebP: RIFF....WEBP
    isWebP = False
    If UBound(allBytes) >= 11 Then
        isWebP = (allBytes(0) = &H52 And allBytes(1) = &H49 And allBytes(2) = &H46 And allBytes(3) = &H46 And _
                  allBytes(8) = &H57 And allBytes(9) = &H45 And allBytes(10) = &H42 And allBytes(11) = &H50)
    End If

    IsValidImageMagicBytes = isJPEG Or isPNG Or isGIF Or isBMP Or isWebP
    On Error GoTo 0
End Function

' --------------------------------------------
' 清理文件名，防止目录遍历
' --------------------------------------------
Function SanitizeFileName(name)
    Dim result

    result = name

    result = Replace(result, "..", "")
    result = Replace(result, "/", "")
    result = Replace(result, "\", "")
    result = Replace(result, ":", "")
    result = Replace(result, "*", "")
    result = Replace(result, "?", "")
    result = Replace(result, Chr(34), "")
    result = Replace(result, "<", "")
    result = Replace(result, ">", "")
    result = Replace(result, "|", "")
    result = Trim(result)

    SanitizeFileName = result
End Function

' --------------------------------------------
' 递归创建目录
' --------------------------------------------
Sub CreateFolderRecursive(fso, folderPath)
    Dim parentPath
    On Error Resume Next

    If Len(folderPath) = 0 Then Exit Sub
    If fso.FolderExists(folderPath) Then Exit Sub

    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        Call CreateFolderRecursive(fso, parentPath)
    End If

    fso.CreateFolder(folderPath)
End Sub

' --------------------------------------------
' 确保上传目录存在
' --------------------------------------------
Function EnsureUploadDir(physicalPath)
    Dim fso, folderPath

    On Error Resume Next
    EnsureUploadDir = False

    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    folderPath = physicalPath
    If fso.FileExists(folderPath) Or InStrRev(folderPath, ".") > InStrRev(folderPath, "\") Then
        folderPath = fso.GetParentFolderName(folderPath)
    End If

    Call CreateFolderRecursive(fso, folderPath)

    If fso.FolderExists(folderPath) Then
        EnsureUploadDir = True
    End If

    Set fso = Nothing
    On Error GoTo 0
End Function
%>

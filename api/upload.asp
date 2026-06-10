<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.ContentType = "application/json"
Response.Charset = "UTF-8"
%><!--#include file="../includes/config.asp"--><!--#include file="../includes/connection.asp"--><!--#include file="../includes/upload_utils.asp"--><%
Response.Clear
Response.ContentType = "application/json"
Response.Charset = "UTF-8"

On Error Resume Next

Dim resultJson, parser, uploadType, saveDir, physicalDir, newFileName, filePath
Dim csrfToken, sessionToken, originalFileName, fileExt
Dim warningMsg

resultJson = ""
warningMsg = ""

' 1. 验证请求方法为 POST
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    resultJson = "{""success"": false, ""error"": ""请求方法必须为 POST""}"
    Response.Write resultJson
    Response.End
End If

' 2. 验证管理员登录状态
If Session("AdminID") = "" Or IsEmpty(Session("AdminID")) Then
    resultJson = "{""success"": false, ""error"": ""未登录或会话已过期""}"
    Response.Write resultJson
    Response.End
End If

' 3. 创建 MultipartParser 实例并解析请求
Set parser = New MultipartParser
parser.Parse()

If parser.ErrorMsg <> "" Then
    resultJson = "{""success"": false, ""error"": """ & Replace(parser.ErrorMsg, Chr(34), Chr(34) & Chr(34)) & """}"
    Response.Write resultJson
    Response.End
End If

' 4. 从解析结果获取 CSRF token 并验证（BinaryRead 后 Request.Form 不可用）
csrfToken = parser.FormField("csrf_token")
sessionToken = Session("CSRFToken")

If sessionToken = "" Or csrfToken = "" Or sessionToken <> csrfToken Then
    resultJson = "{""success"": false, ""error"": ""CSRF 验证失败，请刷新页面重试""}"
    Response.Write resultJson
    Response.End
End If

' 5. 获取 type 表单字段
uploadType = LCase(parser.FormField("type"))

' 6. 验证有文件上传
If Not parser.HasFile Then
    resultJson = "{""success"": false, ""error"": ""没有检测到上传文件""}"
    Response.Write resultJson
    Response.End
End If

originalFileName = parser.FileName

' 7. 验证文件大小不超过 MAX_UPLOAD_SIZE
If parser.FileSize > MAX_UPLOAD_SIZE Then
    resultJson = "{""success"": false, ""error"": ""文件大小超过限制 (最大 " & (MAX_UPLOAD_SIZE \ 1048576) & "MB)""}"
    Response.Write resultJson
    Response.End
End If

' 8. 验证文件类型
If Not IsValidImageType(originalFileName) Then
    resultJson = "{""success"": false, ""error"": ""不支持的文件类型，仅允许 " & ALLOWED_IMAGE_EXTENSIONS & """}"
    Response.Write resultJson
    Response.End
End If

' 9. 验证文件魔数（对于非 SVG 文件）
fileExt = LCase(originalFileName)
If InStrRev(fileExt, ".") > 0 Then
    fileExt = Mid(fileExt, InStrRev(fileExt, "."))
End If

If fileExt <> ".svg" Then
    If Not IsValidImageMagicBytes(parser.FileData) Then
        ' 魔数检查失败时不阻断上传，仅记录警告
        warningMsg = "文件头验证未通过，请确认是有效的图片文件"
    End If
End If

' 10. 根据 type 确定保存目录
Select Case uploadType
    Case "product"
        saveDir = UPLOAD_PATH_PRODUCTS
    Case "note"
        saveDir = UPLOAD_PATH_NOTES
    Case "bottle"
        saveDir = UPLOAD_PATH_BOTTLES
    Case "avatar"
        saveDir = UPLOAD_PATH_AVATARS
    Case Else
        saveDir = UPLOAD_PATH_DEFAULT
End Select

' 11. 使用 EnsureUploadDir 确保目录存在
physicalDir = Server.MapPath(saveDir)
If Right(physicalDir, 1) <> "\" Then
    physicalDir = physicalDir & "\"
End If

If Not EnsureUploadDir(physicalDir) Then
    resultJson = "{""success"": false, ""error"": ""无法创建上传目录""}"
    Response.Write resultJson
    Response.End
End If

' 12. 使用 GenerateUploadFileName 生成唯一文件名
newFileName = GenerateUploadFileName(SanitizeFileName(originalFileName))

' 13. 使用 SaveUploadedFile 保存文件
filePath = physicalDir & newFileName
If SaveUploadedFile(parser.FileData, filePath) Then
    ' 14. 返回 JSON 成功响应
    Dim jsonFileName
    jsonFileName = Replace(originalFileName, Chr(34), "")
    jsonFileName = Replace(jsonFileName, vbCrLf, " ")
    
    resultJson = "{""success"": true, ""url"": """ & saveDir & newFileName & """, ""fileName"": """ & jsonFileName & """, ""fileSize"": " & parser.FileSize
    If warningMsg <> "" Then
        resultJson = resultJson & ", ""warning"": """ & Replace(warningMsg, Chr(34), Chr(34) & Chr(34)) & """"
    End If
    resultJson = resultJson & "}"
Else
    resultJson = "{""success"": false, ""error"": ""文件保存失败，请检查目录权限""}"
End If

If Err.Number <> 0 Then
    resultJson = "{""success"": false, ""error"": ""服务器内部错误: " & Replace(Server.HTMLEncode(Err.Description), Chr(34), "'") & """}"
End If

Response.Write resultJson
On Error GoTo 0
%>

' V18.0 Backup Script
Dim fso, rootDir, bDir, i
Set fso = CreateObject("Scripting.FileSystemObject")

' Path: f:\网站制作\网站\网站二
rootDir = "f:" & ChrW(32593) & ChrW(31449) & ChrW(21046) & ChrW(20316) & _
          "\" & ChrW(32593) & ChrW(31449) & _
          "\" & ChrW(32593) & ChrW(31449) & ChrW(20108)

bDir = rootDir & "\database\backups\V18\code_backup"

Call CreateDirTree(bDir)

Call CopyOne(rootDir & "\includes\payment_config.asp", bDir & "\includes\")
Call CopyOne(rootDir & "\includes\api_auth.asp", bDir & "\includes\")
Call CopyOne(rootDir & "\includes\rate_limiter.asp", bDir & "\includes\")
Call CopyOne(rootDir & "\includes\api_guard.asp", bDir & "\includes\")
Call CopyOne(rootDir & "\includes\cache_v18_ext.asp", bDir & "\includes\")
Call CopyOne(rootDir & "\includes\metrics.asp", bDir & "\includes\")
Call CopyOne(rootDir & "\includes\config.asp", bDir & "\includes\")
Call CopyOne(rootDir & "\css\mobile-first.css", bDir & "\css\")
Call CopyOne(rootDir & "\css\responsive.css", bDir & "\css\")
Call CopyOne(rootDir & "\js\push-manager.js", bDir & "\js\")
Call CopyOne(rootDir & "\js\components\product-card.js", bDir & "\js\components\")
Call CopyOne(rootDir & "\js\components\review-stars.js", bDir & "\js\components\")
Call CopyOne(rootDir & "\js\components\search-autocomplete.js", bDir & "\js\components\")
Call CopyOne(rootDir & "\sw.js", bDir & "\")
Call CopyOne(rootDir & "\offline.html", bDir & "\")
Call CopyOne(rootDir & "\manifest.json", bDir & "\")
Call CopyOne(rootDir & "\checkout.asp", bDir & "\")
Call CopyOne(rootDir & "\index.asp", bDir & "\")
Call CopyOne(rootDir & "\product.asp", bDir & "\")
Call CopyOne(rootDir & "\includes\footer.asp", bDir & "\includes\")
Call CopyOne(rootDir & "\includes\mobile_nav.asp", bDir & "\includes\")
Call CopyOne(rootDir & "\api\notifications_sse.asp", bDir & "\api\")
Call CopyOne(rootDir & "\api\health_check.asp", bDir & "\api\")
Call CopyOne(rootDir & "\admin\analytics\index.asp", bDir & "\admin\analytics\")

Dim scripts : scripts = Array("v18_member_tiers.sql", "v18_points_system.sql", _
    "v18_coupon_system.sql", "v18_flash_group_activities.sql", _
    "v18_subscription.sql", "v18_community_ugc.sql", "v18_perf_indexes.sql")
For i = 0 To 6
    Call CopyOne(rootDir & "\database\" & scripts(i), bDir & "\database\")
Next

Call CopyOne(rootDir & "\docs\" & "V18.0_" & ChrW(23436) & ChrW(25972) & ChrW(24402) & ChrW(26723) & ".md", bDir & "\docs\")
Call CopyOne(rootDir & "\docs\" & "V18.0_" & ChrW(20351) & ChrW(29992) & ChrW(25163) & ChrW(20876) & ".md", bDir & "\docs\")
Call CopyOne(rootDir & "\test_v18_smoke.ps1", bDir & "\")
Call CopyOne(rootDir & "\V18_" & ChrW(20351) & ChrW(29992) & ChrW(27969) & ChrW(31243) & ChrW(25351) & ChrW(21335) & ".html", bDir & "\")

Dim fc : fc = 0
Call CountEm(bDir, fc)
WScript.Echo "V18.0 backup OK: " & fc & " files"

Sub CreateDirTree(base)
    Dim subs : subs = Array("\includes","\css","\js\components","\api","\database","\docs","\admin\analytics")
    If Not fso.FolderExists(base) Then fso.CreateFolder(base)
    For Each s In subs
        If Not fso.FolderExists(base & s) Then fso.CreateFolder(base & s)
    Next
End Sub

Sub CopyOne(src, dst)
    If fso.FileExists(src) Then fso.CopyFile src, dst, True
End Sub

Sub CountEm(fp, c)
    Dim f, sf
    For Each f In fso.GetFolder(fp).Files : c = c + 1 : Next
    For Each sf In fso.GetFolder(fp).SubFolders : Call CountEm(sf.Path, c) : Next
End Sub

const fs = require('fs');
const file = 'f:/网站制作/网站/网站二/admin/purchase/purchase_orders.asp';
let content = fs.readFileSync(file, 'utf8');

// ===== Edit 1: Add ReceivedQty column to details table header in view mode =====
const thPattern = '<th style="text-align:right;">小计</th>';
const viewSectionIdx = content.indexOf('<% If viewMode Then %>');
if (viewSectionIdx >= 0) {
    const detailsStart = content.indexOf(thPattern, viewSectionIdx);
    if (detailsStart >= 0) {
        const before = content.substring(0, detailsStart + thPattern.length);
        const after = content.substring(detailsStart + thPattern.length);
        content = before + '\n                        <th style="text-align:right;">已收货</th>' + after;
        console.log('Edit 1: Added ReceivedQty header column');
    }

    // ===== Edit 2: Add ReceivedQty cell in each detail row in view mode =====
    const rowPattern = 'style="text-align:right;">\u00a5<%= FormatNumber(SafeNum(rsViewDetails("TotalPrice")), 2) %></td>';
    let rowIdx = content.indexOf(rowPattern, viewSectionIdx);
    let count = 0;
    while (rowIdx >= 0 && count < 50) {
        const insertPos = rowIdx + rowPattern.length;
        const before2 = content.substring(0, insertPos);
        const after2 = content.substring(insertPos);
        content = before2 + '\n                        <td style="text-align:right;"><%= SafeNum(rsViewDetails("ReceivedQty")) %></td>' + after2;
        rowIdx = content.indexOf(rowPattern, insertPos + 80);
        count++;
    }
    console.log('Edit 2: Added ReceivedQty cells to ' + count + ' rows');
} else {
    console.log('Edit 1/2: viewMode section not found');
}

// ===== Edit 3: Add operation timeline =====
const statusOpsMarker = "' ========== 状态操作按钮 ==========";
const statusOpsIdx = content.indexOf(statusOpsMarker);
if (statusOpsIdx >= 0) {
    const timelineHTML = `
            ' ========== V12: \u64cd\u4f5c\u65f6\u95f4\u7ebf ==========
            <h4 style="margin:20px 0 15px 0;color:#fff;"><i class="fas fa-history"></i> \u64cd\u4f5c\u65f6\u95f4\u7ebf</h4>
            <div class="timeline">
                <%
                If Not rsStatusLog Is Nothing Then
                    If Not rsStatusLog.EOF Then
                        Dim tlIdx : tlIdx = 0
                        Do While Not rsStatusLog.EOF
                            Dim tlFromStatus : tlFromStatus = CStr(rsStatusLog("FromStatus") & "")
                            Dim tlToStatus : tlToStatus = CStr(rsStatusLog("ToStatus"))
                            Dim tlLogTime : tlLogTime = rsStatusLog("ChangedAt")
                            Dim tlChanger : tlChanger = CStr(rsStatusLog("ChangedBy") & "")
                            Dim tlRemark : tlRemark = CStr(rsStatusLog("Remarks") & "")
                %>
                <div class="timeline-item">
                    <div class="timeline-dot <%= IIf(tlIdx=0, "active", "") %>"></div>
                    <div class="timeline-content">
                        <div class="timeline-header">
                            <span class="status-badge <%= GetStatusClass(tlToStatus) %>"><%= GetStatusName(tlToStatus) %></span>
                            <span class="timeline-time"><% If IsDate(tlLogTime) Then Response.Write FormatDateTime(tlLogTime, 2) & " " & FormatDateTime(tlLogTime, 4) End If %></span>
                        </div>
                        <div class="timeline-desc">
                            <% If tlFromStatus <> "" Then %>
                                <%= GetStatusName(tlFromStatus) %> &rarr; <%= GetStatusName(tlToStatus) %>
                            <% Else %>
                                \u521b\u5efa\u8ba2\u5355\uff08\u521d\u59cb\u72b6\u6001\uff1a<%= GetStatusName(tlToStatus) %>\uff09
                            <% End If %>
                            <% If tlRemark <> "" Then %> &mdash; <em><%= Server.HTMLEncode(tlRemark) %></em><% End If %>
                        </div>
                        <% If tlChanger <> "" Then %>
                        <div class="timeline-actor"><i class="fas fa-user"></i> <%= Server.HTMLEncode(tlChanger) %></div>
                        <% End If %>
                    </div>
                </div>
                <%
                            tlIdx = tlIdx + 1
                            rsStatusLog.MoveNext
                        Loop
                        rsStatusLog.Close
                        Set rsStatusLog = Nothing
                    Else
                        rsStatusLog.Close
                        Set rsStatusLog = Nothing
                %>
                <div class="timeline-item">
                    <div class="timeline-dot"></div>
                    <div class="timeline-content">
                        <div class="timeline-desc" style="color:#666;">\u6682\u65e0\u72b6\u6001\u53d8\u66f4\u8bb0\u5f55</div>
                    </div>
                </div>
                <%
                    End If
                Else
                %>
                <div class="timeline-item">
                    <div class="timeline-dot"></div>
                    <div class="timeline-content">
                        <div class="timeline-desc" style="color:#666;">\u6682\u65e0\u72b6\u6001\u53d8\u66f4\u8bb0\u5f55</div>
                    </div>
                </div>
                <% End If %>
            </div>`;

    const before3 = content.substring(0, statusOpsIdx);
    const after3 = content.substring(statusOpsIdx);
    content = before3 + timelineHTML + '\n' + after3;
    console.log('Edit 3: Added operation timeline');
} else {
    console.log('Edit 3: status ops marker not found');
}

// ===== Edit 4: Add timeline CSS styles =====
const styleEnd = content.indexOf('</style>');
if (styleEnd >= 0) {
    const timelineCSS = `
        /* V12: 操作时间线 */
        .timeline {
            position: relative;
            padding-left: 30px;
        }
        .timeline::before {
            content: '';
            position: absolute;
            left: 10px;
            top: 0;
            bottom: 0;
            width: 2px;
            background: rgba(255,255,255,0.08);
        }
        .timeline-item {
            position: relative;
            margin-bottom: 20px;
        }
        .timeline-item:last-child {
            margin-bottom: 0;
        }
        .timeline-dot {
            position: absolute;
            left: -24px;
            top: 4px;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background: rgba(255,255,255,0.15);
            border: 2px solid rgba(255,255,255,0.1);
        }
        .timeline-dot.active {
            background: #FF9800;
            border-color: #FF9800;
            box-shadow: 0 0 8px rgba(255,152,0,0.4);
        }
        .timeline-content {
            background: rgba(255,255,255,0.02);
            padding: 12px 15px;
            border-radius: 8px;
            border: 1px solid rgba(255,255,255,0.04);
        }
        .timeline-header {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 6px;
        }
        .timeline-time {
            font-size: 11px;
            color: #888;
        }
        .timeline-desc {
            font-size: 13px;
            color: #ccc;
        }
        .timeline-actor {
            font-size: 11px;
            color: #888;
            margin-top: 4px;
        }`;

    const before4 = content.substring(0, styleEnd);
    const after4 = content.substring(styleEnd);
    content = before4 + timelineCSS + '\n' + after4;
    console.log('Edit 4: Added timeline CSS');
}

fs.writeFileSync(file, content, 'utf8');
console.log('All edits applied successfully');

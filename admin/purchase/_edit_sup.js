const fs = require('fs');
const file = 'f:/网站制作/网站/网站二/admin/purchase/supplier_management.asp';
let content = fs.readFileSync(file, 'utf8');

// ===== Edit 1: Add supplier rating, recent orders, and active contracts queries =====
const detailMarker = "detailExists = True";
const detailIdx = content.indexOf(detailMarker);
if (detailIdx >= 0) {
    // Find the closing End If after this block
    const afterDetail = content.indexOf("End If", detailIdx + detailMarker.length);
    const nextEndIf = content.indexOf("End If", afterDetail + 7);
    
    // Insert the new queries after the detail loading block
    const insertPos = nextEndIf + 7; // after "End If"
    const queriesCode = `
' ========== V12: \u83b7\u53d6\u4f9b\u5e94\u5546\u5408\u4f5c\u5386\u53f2\u4e0e\u8bc4\u5206 ==========
Dim rsSupplierOrders, supplierRatingAvg, supplierEvalCount, supplierActiveContracts
If detailExists Then
    ' \u6700\u8fd15\u7b14\u91c7\u8d2d\u8ba2\u5355
    Set rsSupplierOrders = ExecuteQuery("SELECT TOP 5 PurchaseNo, OrderDate, CAST(ISNULL(TotalAmount,0) AS FLOAT) as TotalAmount, Status FROM PurchaseOrders WHERE SupplierID=" & CInt(detailSupplierId) & " ORDER BY OrderDate DESC")
    
    ' \u4f9b\u5e94\u5546\u7efc\u5408\u8bc4\u5206
    Dim ratingRS : Set ratingRS = ExecuteQuery("SELECT AVG(CAST(OverallScore AS FLOAT)) as AvgScore, COUNT(*) as EvalCount FROM SupplierEvaluations WHERE SupplierID=" & CInt(detailSupplierId))
    If Not ratingRS Is Nothing Then
        If Not ratingRS.EOF Then
            supplierRatingAvg = SafeNum(ratingRS("AvgScore"))
            supplierEvalCount = SafeNum(ratingRS("EvalCount"))
        End If
        ratingRS.Close : Set ratingRS = Nothing
    End If
    
    ' \u6d3b\u8dc3\u5408\u540c\u6570
    supplierActiveContracts = SafeNum(GetScalar("SELECT COUNT(*) FROM SupplierContracts WHERE SupplierID=" & CInt(detailSupplierId) & " AND Status='Active'"))
End If`;

    const before = content.substring(0, insertPos);
    const after = content.substring(insertPos);
    content = before + queriesCode + after;
    console.log('Edit 1: Added supplier history queries');
}

// ===== Edit 2: Enhance the detail view with cooperation history, rating, and active contracts =====
// Find the "业务统计" section and add more stats
const statsMarker = '\u4e1a\u52a1\u7edf\u8ba1'; // 业务统计
const statsIdx = content.indexOf(statsMarker);
if (statsIdx >= 0) {
    // Find the closing </div> of the detail-stats grid
    const gridEnd = content.indexOf('</div>', content.indexOf('detail-stats', statsIdx));
    const sectionEnd = content.indexOf('</div>', gridEnd + 6); // end of detail-section
    
    // Add more stat items to the existing stats grid
    const statsGridStart = content.indexOf('detail-stats', statsIdx);
    if (statsGridStart >= 0) {
        const gridContentStart = content.indexOf('>', statsGridStart) + 1;
        
        // Add rating and active contracts stats
        const newStats = `
                    <div class="detail-stat-item">
                        <div class="detail-stat-value"><%= supplierEvalCount %></div>
                        <div class="detail-stat-label">\u8bc4\u4f30\u6b21\u6570</div>
                    </div>
                    <div class="detail-stat-item">
                        <div class="detail-stat-value"><%= FormatNumber(supplierRatingAvg, 1) %></div>
                        <div class="detail-stat-label">\u7efc\u5408\u8bc4\u5206</div>
                    </div>
                    <div class="detail-stat-item">
                        <div class="detail-stat-value"><%= supplierActiveContracts %></div>
                        <div class="detail-stat-label">\u6d3b\u8dc3\u5408\u540c</div>
                    </div>`;
        
        // Find the closing of the existing stats grid (after the last stat item)
        const lastStatItem = content.lastIndexOf('detail-stat-item', gridEnd);
        const lastItemClose = content.indexOf('</div>', lastStatItem) + 6;
        
        const before2 = content.substring(0, lastItemClose);
        const after2 = content.substring(lastItemClose);
        content = before2 + newStats + after2;
        console.log('Edit 2: Added rating and contract stats');
    }
    
    // ===== Edit 3: Add recent purchase orders section after stats =====
    // Find end of detail-section containing stats
    const statsSectionEnd = content.indexOf('</div>', gridEnd + 6) + 6;
    
    const ordersHTML = `
            
            <div class="detail-section">
                <h4><i class="fas fa-history"></i> \u5408\u4f5c\u5386\u53f2</h4>
                <%
                If Not rsSupplierOrders Is Nothing Then
                    If Not rsSupplierOrders.EOF Then
                %>
                <table class="data-table" style="font-size:12px;">
                    <thead>
                        <tr>
                            <th>\u8ba2\u5355\u53f7</th>
                            <th>\u65e5\u671f</th>
                            <th style="text-align:right;">\u91d1\u989d</th>
                            <th>\u72b6\u6001</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        Do While Not rsSupplierOrders.EOF
                        %>
                        <tr>
                            <td><a href="purchase_orders.asp?view=<%= rsSupplierOrders("PurchaseNo") %>" style="color:#FF9800;"><%= Server.HTMLEncode(rsSupplierOrders("PurchaseNo") & "") %></a></td>
                            <td><% If IsDate(rsSupplierOrders("OrderDate")) Then Response.Write FormatDateTime(rsSupplierOrders("OrderDate"), 2) %></td>
                            <td style="text-align:right;">\u00a5<%= FormatNumber(SafeNum(rsSupplierOrders("TotalAmount")), 2) %></td>
                            <td><span class="status-badge status-<%= LCase(rsSupplierOrders("Status") & "") %>"><%= rsSupplierOrders("Status") & "" %></span></td>
                        </tr>
                        <%
                            rsSupplierOrders.MoveNext
                        Loop
                        rsSupplierOrders.Close : Set rsSupplierOrders = Nothing
                        %>
                    </tbody>
                </table>
                <%
                    Else
                        rsSupplierOrders.Close : Set rsSupplierOrders = Nothing
                %>
                <div style="color:#666;padding:10px 0;">\u6682\u65e0\u91c7\u8d2d\u8bb0\u5f55</div>
                <%
                    End If
                Else
                %>
                <div style="color:#666;padding:10px 0;">\u6682\u65e0\u91c7\u8d2d\u8bb0\u5f55</div>
                <% End If %>
            </div>`;
    
    const before3 = content.substring(0, statsSectionEnd);
    const after3 = content.substring(statsSectionEnd);
    content = before3 + ordersHTML + after3;
    console.log('Edit 3: Added recent purchase orders section');
}

fs.writeFileSync(file, content, 'utf8');
console.log('All edits applied successfully');

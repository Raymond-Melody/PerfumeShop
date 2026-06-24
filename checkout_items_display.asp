<!-- ============================================
     V14.6 结算页 - 购物车商品展示模板
     从 checkout.asp 提取
     ============================================ -->
        <div class="checkout-content">
            <div class="checkout-items">
                <h3>订单商品</h3>
                
                <%
                Set rsCart = ExecuteQuery("SELECT c.*, p.ProductName, p.ImageURL, p.EngravingPrice, p.ProductType, " & _
                    "tn.NoteName AS TopNoteName, mn.NoteName AS MiddleNoteName, bn.NoteName AS BaseNoteName, " & _
                    "v.VolumeName, v.VolumeML, b.BottleName, " & _
                    "c.Quantity * c.UnitPrice AS SubTotal " & _
                    "FROM ((((((Cart c " & _
                    "LEFT JOIN Products p ON c.ProductID = p.ProductID) " & _
                    "LEFT JOIN FragranceNotes tn ON c.TopNoteID = tn.NoteID) " & _
                    "LEFT JOIN FragranceNotes mn ON c.MiddleNoteID = mn.NoteID) " & _
                    "LEFT JOIN FragranceNotes bn ON c.BaseNoteID = bn.NoteID) " & _
                    "LEFT JOIN Volumes v ON c.VolumeID = v.VolumeID) " & _
                    "LEFT JOIN BottleStyles b ON c.BottleID = b.BottleID) " & _
                    "WHERE " & whereClause & " ORDER BY c.CreatedAt DESC")
                
                If Not rsCart Is Nothing Then
                    Do While Not rsCart.EOF
                %>
                <div class="checkout-item">
                    <div class="item-image">
                        <img src="<%= rsCart("ImageURL") %>" alt="<%= HTMLEncode(rsCart("ProductName")) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                    </div>
                    <div class="item-details">
                        <h4><%= HTMLEncode(rsCart("ProductName")) %></h4>
                        <div class="item-attributes">
                            <%
                            Dim chkProductType, chkProductTypeLC, chkCartId
                            chkProductType = rsCart("ProductType") & ""
                            chkProductTypeLC = LCase(chkProductType)
                            chkCartId = rsCart("CartID")
                            
                            Dim rsChkNotes, chkTopList, chkMidList, chkBaseList, chkNoteType, chkNoteName, chkPercent
                            chkTopList = "": chkMidList = "": chkBaseList = ""
                            If chkProductTypeLC = "custom" Then
                                Set rsChkNotes = ExecuteQuery("SELECT n.NoteName, s.Percentage, s.NoteType FROM CartNoteSelections s INNER JOIN FragranceNotes n ON s.NoteID = n.NoteID WHERE s.CartID = " & chkCartId)
                                If Not rsChkNotes Is Nothing Then
                                    Do While Not rsChkNotes.EOF
                                        chkNoteType = Trim(rsChkNotes("NoteType") & "")
                                        chkNoteName = HTMLEncode(rsChkNotes("NoteName") & "")
                                        chkPercent = rsChkNotes("Percentage")
                                        If chkNoteType = "前调" Then
                                            If chkTopList <> "" Then chkTopList = chkTopList & ", "
                                            chkTopList = chkTopList & chkNoteName & " (" & chkPercent & "%)"
                                        ElseIf chkNoteType = "中调" Then
                                            If chkMidList <> "" Then chkMidList = chkMidList & ", "
                                            chkMidList = chkMidList & chkNoteName & " (" & chkPercent & "%)"
                                        ElseIf chkNoteType = "后调" Then
                                            If chkBaseList <> "" Then chkBaseList = chkBaseList & ", "
                                            chkBaseList = chkBaseList & chkNoteName & " (" & chkPercent & "%)"
                                        End If
                                        rsChkNotes.MoveNext
                                    Loop
                                    rsChkNotes.Close
                                    Set rsChkNotes = Nothing
                                End If
                            End If
                            
                            If chkProductTypeLC = "custom" And chkTopList <> "" Then %>
                            <span><i class="fas fa-wind"></i> 前调: <%= chkTopList %></span>
                            <% ElseIf chkProductTypeLC = "custom" And Not IsNull(rsCart("TopNoteName")) Then %>
                            <span><i class="fas fa-wind"></i> 前调: <%= HTMLEncode(rsCart("TopNoteName")) %></span>
                            <% End If %>
                            <% If chkProductTypeLC = "custom" And chkMidList <> "" Then %>
                            <span><i class="fas fa-heart"></i> 中调: <%= chkMidList %></span>
                            <% ElseIf chkProductTypeLC = "custom" And Not IsNull(rsCart("MiddleNoteName")) Then %>
                            <span><i class="fas fa-heart"></i> 中调: <%= HTMLEncode(rsCart("MiddleNoteName")) %></span>
                            <% End If %>
                            <% If chkProductTypeLC = "custom" And chkBaseList <> "" Then %>
                            <span><i class="fas fa-moon"></i> 后调: <%= chkBaseList %></span>
                            <% ElseIf chkProductTypeLC = "custom" And Not IsNull(rsCart("BaseNoteName")) Then %>
                            <span><i class="fas fa-moon"></i> 后调: <%= HTMLEncode(rsCart("BaseNoteName")) %></span>
                            <% End If %>
                            <% If Not IsNull(rsCart("VolumeName")) Then %>
                            <span><i class="fas fa-tint"></i> 容量: <%= rsCart("VolumeML") %>ml (<%= HTMLEncode(rsCart("VolumeName")) %>)</span>
                            <% End If %>
                            <% If Not IsNull(rsCart("BottleName")) Then %>
                            <span><i class="fas fa-wine-bottle"></i> 瓶身: <%= HTMLEncode(rsCart("BottleName")) %></span>
                            <% End If %>
                            <% If Not IsNull(rsCart("CustomLabel")) And rsCart("CustomLabel") <> "" Then %>
                            <span><i class="fas fa-pen-fancy"></i> 刻字: <%= HTMLEncode(rsCart("CustomLabel")) %></span>
                            <% End If %>
                            <% 
                            ' 显示刻字费用
                            Dim checkoutItemEngravingPrice
                            checkoutItemEngravingPrice = 0
                            On Error Resume Next
                            checkoutItemEngravingPrice = CDbl(rsCart("EngravingPrice"))
                            If Err.Number <> 0 Then checkoutItemEngravingPrice = 0
                            On Error GoTo 0
                            If Not IsNull(rsCart("CustomLabel")) And rsCart("CustomLabel") <> "" And checkoutItemEngravingPrice > 0 Then 
                            %>
                            <span style="color:#e91e63;"><i class="fas fa-tag"></i> 刻字费用: <%= FormatMoney(checkoutItemEngravingPrice) %></span>
                            <% End If %>
                        </div>
                    </div>
                    <div class="item-quantity">
                        × <%= rsCart("Quantity") %>
                    </div>
                    <div class="item-price">
                        <%= FormatMoney(rsCart("SubTotal")) %>
                    </div>
                </div>
                <%
                    rsCart.MoveNext
                    Loop
                    rsCart.Close
                    Set rsCart = Nothing
                End If
                %>
            </div>
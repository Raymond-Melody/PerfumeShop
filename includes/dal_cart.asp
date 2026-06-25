<%
' ============================================
' V17.0 DAL - 购物车数据访问层
' 依赖: dal.asp, connection.asp
' 用法: <!--#include file="dal_cart.asp"-->
' 涵盖: Cart, CartNoteSelections
' ============================================

' ============================================
' 获取用户购物车（含产品、香调、容量、瓶身信息）
' ============================================
Function DAL_Cart_GetByUser(userId)
    Dim sql, params(0)
    sql = "SELECT c.*, p.ProductName, p.ImageURL, p.EngravingPrice, p.ProductType, " & _
          "tn.NoteName AS TopNoteName, mn.NoteName AS MiddleNoteName, bn.NoteName AS BaseNoteName, " & _
          "v.VolumeName, v.VolumeML, b.BottleName " & _
          "FROM Cart c " & _
          "LEFT JOIN Products p ON c.ProductID = p.ProductID " & _
          "LEFT JOIN FragranceNotes tn ON c.TopNoteID = tn.NoteID " & _
          "LEFT JOIN FragranceNotes mn ON c.MiddleNoteID = mn.NoteID " & _
          "LEFT JOIN FragranceNotes bn ON c.BaseNoteID = bn.NoteID " & _
          "LEFT JOIN Volumes v ON c.VolumeID = v.VolumeID " & _
          "LEFT JOIN BottleStyles b ON c.BottleID = b.BottleID " & _
          "WHERE c.UserID = @UserID ORDER BY c.CreatedAt DESC"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    Set DAL_Cart_GetByUser = DAL_GetList(sql, params)
End Function

' ============================================
' 获取匿名购物车（基于SessionID）
' ============================================
Function DAL_Cart_GetBySession(sessionId)
    Dim sql, params(0)
    sql = "SELECT c.*, p.ProductName, p.ImageURL, p.EngravingPrice, p.ProductType, " & _
          "tn.NoteName AS TopNoteName, mn.NoteName AS MiddleNoteName, bn.NoteName AS BaseNoteName, " & _
          "v.VolumeName, v.VolumeML, b.BottleName " & _
          "FROM Cart c " & _
          "LEFT JOIN Products p ON c.ProductID = p.ProductID " & _
          "LEFT JOIN FragranceNotes tn ON c.TopNoteID = tn.NoteID " & _
          "LEFT JOIN FragranceNotes mn ON c.MiddleNoteID = mn.NoteID " & _
          "LEFT JOIN FragranceNotes bn ON c.BaseNoteID = bn.NoteID " & _
          "LEFT JOIN Volumes v ON c.VolumeID = v.VolumeID " & _
          "LEFT JOIN BottleStyles b ON c.BottleID = b.BottleID " & _
          "WHERE c.SessionID = @SessionID ORDER BY c.CreatedAt DESC"
    params(0) = Array("@SessionID", DAL_adVarChar, 100, sessionId)
    Set DAL_Cart_GetBySession = DAL_GetList(sql, params)
End Function

' ============================================
' 获取购物车商品数量
' ============================================
Function DAL_Cart_GetCount(userId, sessionId)
    Dim sql
    If userId <> "" Then
        sql = "SELECT COUNT(*) FROM Cart WHERE UserID=" & CLng(userId)
    Else
        sql = "SELECT COUNT(*) FROM Cart WHERE SessionID='" & Replace(sessionId, "'", "''") & "'"
    End If
    DAL_Cart_GetCount = CLng(DAL_GetScalar(sql, Null, 0))
End Function

' ============================================
' 通过CartID获取购物车项
' ============================================
Function DAL_Cart_GetByID(cartId)
    Dim sql, params(0)
    sql = "SELECT * FROM Cart WHERE CartID=@CartID"
    params(0) = Array("@CartID", DAL_adInteger, 0, CLng(cartId))
    Set DAL_Cart_GetByID = DAL_GetRow(sql, params)
End Function

' ============================================
' 添加商品到购物车
' ============================================
Function DAL_Cart_Add(userId, sessionId, productId, volumeId, bottleId, topNoteId, middleNoteId, baseNoteId, quantity, unitPrice, customLabel)
    Dim params(9)
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    params(1) = Array("@SessionID", DAL_adVarChar, 100, Left(sessionId, 100))
    params(2) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    params(3) = Array("@VolumeID", DAL_adInteger, 0, CLng(volumeId))
    params(4) = Array("@BottleID", DAL_adInteger, 0, CLng(bottleId))
    params(5) = Array("@TopNoteID", DAL_adInteger, 0, CLng(topNoteId))
    params(6) = Array("@MiddleNoteID", DAL_adInteger, 0, CLng(middleNoteId))
    params(7) = Array("@BaseNoteID", DAL_adInteger, 0, CLng(baseNoteId))
    params(8) = Array("@Quantity", DAL_adInteger, 0, CInt(quantity))
    params(9) = Array("@UnitPrice", DAL_adCurrency, 0, CDbl(unitPrice))
    
    Dim sql
    sql = "INSERT INTO Cart (UserID, SessionID, ProductID, VolumeID, BottleID, " & _
          "TopNoteID, MiddleNoteID, BaseNoteID, Quantity, UnitPrice, CustomLabel) " & _
          "VALUES (@UserID, @SessionID, @ProductID, @VolumeID, @BottleID, " & _
          "@TopNoteID, @MiddleNoteID, @BaseNoteID, @Quantity, @UnitPrice, @CustomLabel); SELECT SCOPE_IDENTITY()"
    
    ReDim Preserve params(10)
    params(10) = Array("@CustomLabel", DAL_adVarChar, 200, Left(customLabel, 200))
    
    DAL_Cart_Add = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 更新购物车数量
' ============================================
Function DAL_Cart_UpdateQuantity(cartId, quantity)
    Dim sql, params(1)
    sql = "UPDATE Cart SET Quantity=@Quantity WHERE CartID=@CartID"
    params(0) = Array("@Quantity", DAL_adInteger, 0, CInt(quantity))
    params(1) = Array("@CartID", DAL_adInteger, 0, CLng(cartId))
    DAL_Cart_UpdateQuantity = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 删除购物车项
' ============================================
Function DAL_Cart_Remove(cartId)
    Dim sql, params(0)
    
    ' 先删除关联的香调选择
    Dim delNotes(0)
    delNotes(0) = Array("@CartID", DAL_adInteger, 0, CLng(cartId))
    DAL_Execute "DELETE FROM CartNoteSelections WHERE CartID=@CartID", delNotes
    
    ' 再删除购物车项
    sql = "DELETE FROM Cart WHERE CartID=@CartID"
    params(0) = Array("@CartID", DAL_adInteger, 0, CLng(cartId))
    DAL_Cart_Remove = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 清空用户购物车
' ============================================
Function DAL_Cart_ClearByUser(userId)
    Dim sql, params(0)
    sql = "DELETE FROM Cart WHERE UserID=@UserID"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    DAL_Cart_ClearByUser = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 清空匿名购物车
' ============================================
Function DAL_Cart_ClearBySession(sessionId)
    Dim sql, params(0)
    sql = "DELETE FROM Cart WHERE SessionID=@SessionID"
    params(0) = Array("@SessionID", DAL_adVarChar, 100, sessionId)
    DAL_Cart_ClearBySession = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 获取购物车香调选择
' ============================================
Function DAL_Cart_GetNoteSelections(cartId)
    Dim sql, params(0)
    sql = "SELECT s.*, n.NoteName, n.NoteType FROM CartNoteSelections s " & _
          "INNER JOIN FragranceNotes n ON s.NoteID = n.NoteID WHERE s.CartID=@CartID"
    params(0) = Array("@CartID", DAL_adInteger, 0, CLng(cartId))
    Set DAL_Cart_GetNoteSelections = DAL_GetList(sql, params)
End Function

' ============================================
' 合并匿名购物车到用户
' ============================================
Sub DAL_Cart_MergeSessionToUser(sessionId, userId)
    Dim sql, params(1)
    sql = "UPDATE Cart SET UserID=@UserID, SessionID=NULL " & _
          "WHERE SessionID=@SessionID AND (UserID IS NULL OR UserID = 0)"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    params(1) = Array("@SessionID", DAL_adVarChar, 100, sessionId)
    DAL_Execute sql, params
End Sub
%>

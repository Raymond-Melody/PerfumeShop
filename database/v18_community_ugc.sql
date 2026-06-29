-- ============================================
-- V18 会员社区与 UGC (Community & UGC)
-- 创建时间: 2026-06-30
-- ============================================

-- ============================================
-- 产品评价表 (ProductReviews)
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[ProductReviews] (
        ReviewID INT IDENTITY(1,1) PRIMARY KEY,
        ProductID INT NOT NULL,
        UserID INT NOT NULL,
        OrderID INT NULL,                           -- 关联订单（验证购买）
        Rating INT NOT NULL CHECK (Rating BETWEEN 1 AND 5),
        Title NVARCHAR(100) NULL,
        Content NVARCHAR(2000) NOT NULL,
        IsVerifiedPurchase BIT NOT NULL DEFAULT 0,  -- 是否已验证购买
        AIFeelingSummary NVARCHAR(500) NULL,        -- AI情感摘要
        LikeCount INT NOT NULL DEFAULT 0,
        IsActive BIT NOT NULL DEFAULT 1,
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        UpdatedAt DATETIME NULL,
        FOREIGN KEY (ProductID) REFERENCES [Products](ProductID),
        FOREIGN KEY (UserID) REFERENCES [Users](UserID)
    );

    CREATE INDEX IX_ProductReviews_Product ON [ProductReviews](ProductID, IsActive, CreatedAt DESC);
    CREATE INDEX IX_ProductReviews_User ON [ProductReviews](UserID, CreatedAt DESC);
END
GO

-- ============================================
-- 评价图片表 (ReviewImages)
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ReviewImages]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[ReviewImages] (
        ImageID INT IDENTITY(1,1) PRIMARY KEY,
        ReviewID INT NOT NULL,
        ImageURL VARCHAR(500) NOT NULL,
        SortOrder INT NOT NULL DEFAULT 0,
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        FOREIGN KEY (ReviewID) REFERENCES [ProductReviews](ReviewID) ON DELETE CASCADE
    );

    CREATE INDEX IX_ReviewImages_Review ON [ReviewImages](ReviewID);
END
GO

-- ============================================
-- 社区帖子表 (CommunityPosts)
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CommunityPosts]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[CommunityPosts] (
        PostID INT IDENTITY(1,1) PRIMARY KEY,
        UserID INT NOT NULL,
        Title NVARCHAR(200) NOT NULL,
        Content NVARCHAR(4000) NOT NULL,
        PostType VARCHAR(20) NOT NULL DEFAULT 'discussion',  -- recipe/review/discussion
        FragranceNotes NVARCHAR(500) NULL,      -- JSON: 前中后调
        Tags NVARCHAR(300) NULL,                -- 逗号分隔标签
        IsPublic BIT NOT NULL DEFAULT 1,
        LikeCount INT NOT NULL DEFAULT 0,
        CommentCount INT NOT NULL DEFAULT 0,
        ViewCount INT NOT NULL DEFAULT 0,
        IsPinned BIT NOT NULL DEFAULT 0,
        IsActive BIT NOT NULL DEFAULT 1,
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        UpdatedAt DATETIME NULL,
        FOREIGN KEY (UserID) REFERENCES [Users](UserID)
    );

    CREATE INDEX IX_CommunityPosts_Type ON [CommunityPosts](PostType, IsActive, CreatedAt DESC);
    CREATE INDEX IX_CommunityPosts_User ON [CommunityPosts](UserID, CreatedAt DESC);
    CREATE INDEX IX_CommunityPosts_Hot ON [CommunityPosts](IsActive, IsPinned DESC, LikeCount DESC, CreatedAt DESC);
END
GO

-- ============================================
-- 帖子评论表 (PostComments)
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PostComments]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[PostComments] (
        CommentID INT IDENTITY(1,1) PRIMARY KEY,
        PostID INT NOT NULL,
        UserID INT NOT NULL,
        ParentCommentID INT NULL,               -- 父评论ID（嵌套回复）
        Content NVARCHAR(1000) NOT NULL,
        IsActive BIT NOT NULL DEFAULT 1,
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        FOREIGN KEY (PostID) REFERENCES [CommunityPosts](PostID) ON DELETE CASCADE,
        FOREIGN KEY (UserID) REFERENCES [Users](UserID)
    );

    CREATE INDEX IX_PostComments_Post ON [PostComments](PostID, IsActive, CreatedAt ASC);
END
GO

-- ============================================
-- 帖子点赞表 (PostLikes)
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PostLikes]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[PostLikes] (
        LikeID INT IDENTITY(1,1) PRIMARY KEY,
        PostID INT NOT NULL,
        UserID INT NOT NULL,
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        FOREIGN KEY (PostID) REFERENCES [CommunityPosts](PostID) ON DELETE CASCADE,
        FOREIGN KEY (UserID) REFERENCES [Users](UserID),
        CONSTRAINT UQ_PostLikes UNIQUE (PostID, UserID)
    );
END
GO

-- ============================================
-- 评价点赞表 (ReviewLikes)
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ReviewLikes]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[ReviewLikes] (
        LikeID INT IDENTITY(1,1) PRIMARY KEY,
        ReviewID INT NOT NULL,
        UserID INT NOT NULL,
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        FOREIGN KEY (ReviewID) REFERENCES [ProductReviews](ReviewID) ON DELETE CASCADE,
        FOREIGN KEY (UserID) REFERENCES [Users](UserID),
        CONSTRAINT UQ_ReviewLikes UNIQUE (ReviewID, UserID)
    );
END
GO

-- ============================================
-- 香调配方分享表 (FragranceNotes)
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FragranceNotes]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[FragranceNotes] (
        NoteID INT IDENTITY(1,1) PRIMARY KEY,
        UserID INT NOT NULL,
        Name NVARCHAR(100) NOT NULL,
        TopNotes NVARCHAR(300) NOT NULL,        -- 前调
        MiddleNotes NVARCHAR(300) NOT NULL,      -- 中调
        BaseNotes NVARCHAR(300) NOT NULL,        -- 后调
        Description NVARCHAR(1000) NULL,         -- 配方描述
        IsPublic BIT NOT NULL DEFAULT 1,
        LikeCount INT NOT NULL DEFAULT 0,
        ViewCount INT NOT NULL DEFAULT 0,
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        FOREIGN KEY (UserID) REFERENCES [Users](UserID)
    );

    CREATE INDEX IX_FragranceNotes_Public ON [FragranceNotes](IsPublic, LikeCount DESC, CreatedAt DESC);
    CREATE INDEX IX_FragranceNotes_User ON [FragranceNotes](UserID);
END
GO

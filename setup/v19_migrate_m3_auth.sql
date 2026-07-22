-- ============================================
-- V19 M3-A: 密码重置令牌表 (幂等)
-- 对齐 V18 Users.ResetToken + ResetTokenExpiry 独立化
-- ============================================

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PasswordResetTokens')
BEGIN
    CREATE TABLE PasswordResetTokens (
        TokenId       INT IDENTITY(1,1) PRIMARY KEY,
        UserId        INT NOT NULL,
        Token         NVARCHAR(128) NOT NULL,
        TokenHash     NVARCHAR(256) NOT NULL,
        ExpiresAt     DATETIME2 NOT NULL,
        IsUsed        BIT NOT NULL DEFAULT 0,
        CreatedAt     DATETIME2 NOT NULL DEFAULT GETDATE(),
        UsedAt        DATETIME2 NULL,
        CONSTRAINT FK_PasswordResetTokens_Users FOREIGN KEY (UserId) REFERENCES Users(UserId)
    );

    CREATE INDEX IX_PasswordResetTokens_TokenHash ON PasswordResetTokens(TokenHash);
    CREATE INDEX IX_PasswordResetTokens_UserId    ON PasswordResetTokens(UserId);
    PRINT 'PasswordResetTokens table created.';
END
ELSE
    PRINT 'PasswordResetTokens table already exists.';
GO

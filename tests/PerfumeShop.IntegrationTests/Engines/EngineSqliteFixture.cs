using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;
using PerfumeShop.Data.Models;

namespace PerfumeShop.IntegrationTests.Engines;

/// <summary>
/// SQLite 兼容的测试 DbContext
/// 继承 EngineTestContext 并在 model 构建后移除 SQL Server 特有的 HasDefaultValueSql("(getdate())")
/// </summary>
public class EngineSqliteContext : EngineTestContext
{
    public EngineSqliteContext(DbContextOptions<PerfumeShopContext> options) : base(options) { }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // 移除 SQL Server 特有的 HasDefaultValueSql("(getdate())") — SQLite 不支持非常量默认值
        foreach (var entityType in modelBuilder.Model.GetEntityTypes())
        {
            foreach (var property in entityType.GetProperties())
            {
                var defaultSql = property.GetDefaultValueSql();
                if (!string.IsNullOrEmpty(defaultSql) &&
                    defaultSql.Contains("getdate", StringComparison.OrdinalIgnoreCase))
                {
                    property.SetDefaultValueSql(null);
                }
            }
        }
    }
}

/// <summary>
/// SQLite 测试夹具 — 每个测试独立 SQLite 内存数据库
/// </summary>
public class SqliteTestBase : IDisposable
{
    protected readonly EngineSqliteContext Db;
    protected readonly SqliteConnection Connection;

    public SqliteTestBase()
    {
        Connection = new SqliteConnection("DataSource=:memory:");
        Connection.Open();

        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseSqlite(Connection)
            .Options;

        Db = new EngineSqliteContext(options);
        Db.Database.EnsureCreated();
    }

    public void Dispose()
    {
        Db.Dispose();
        Connection.Close();
        Connection.Dispose();
    }
}

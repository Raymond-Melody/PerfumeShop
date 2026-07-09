using System.Linq.Expressions;

namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 泛型仓储接口 — 定义标准 CRUD 操作，与 EF Core 解耦
/// </summary>
/// <typeparam name="T">实体类型 (class)</typeparam>
public interface IRepository<T> where T : class
{
    // ========== 查询 ==========

    /// <summary>按主键 ID 获取实体</summary>
    Task<T?> GetByIdAsync(int id, CancellationToken ct = default);

    /// <summary>获取所有实体</summary>
    Task<IEnumerable<T>> GetAllAsync(CancellationToken ct = default);

    /// <summary>按条件查找实体集合</summary>
    Task<IEnumerable<T>> FindAsync(Expression<Func<T, bool>> predicate, CancellationToken ct = default);

    /// <summary>按条件查找单个实体，无匹配返回 null</summary>
    Task<T?> SingleOrDefaultAsync(Expression<Func<T, bool>> predicate, CancellationToken ct = default);

    /// <summary>按条件查找第一个实体，无匹配返回 null</summary>
    Task<T?> FirstOrDefaultAsync(Expression<Func<T, bool>> predicate, CancellationToken ct = default);

    /// <summary>判断是否存在满足条件的实体</summary>
    Task<bool> AnyAsync(Expression<Func<T, bool>> predicate, CancellationToken ct = default);

    /// <summary>统计满足条件的实体数量</summary>
    Task<int> CountAsync(Expression<Func<T, bool>>? predicate = null, CancellationToken ct = default);

    /// <summary>分页查询</summary>
    Task<(IEnumerable<T> Items, int TotalCount)> GetPagedAsync(
        int page, int pageSize,
        Expression<Func<T, bool>>? predicate = null,
        Func<IQueryable<T>, IOrderedQueryable<T>>? orderBy = null,
        CancellationToken ct = default);

    // ========== 新增 ==========

    /// <summary>添加实体 (返回带 Tracked 状态的实体)</summary>
    Task<T> AddAsync(T entity, CancellationToken ct = default);

    /// <summary>批量添加</summary>
    Task AddRangeAsync(IEnumerable<T> entities, CancellationToken ct = default);

    // ========== 更新 ==========

    /// <summary>标记实体为已修改</summary>
    void Update(T entity);

    /// <summary>批量标记修改</summary>
    void UpdateRange(IEnumerable<T> entities);

    // ========== 删除 ==========

    /// <summary>标记实体为待删除</summary>
    void Delete(T entity);

    /// <summary>批量标记删除</summary>
    void DeleteRange(IEnumerable<T> entities);

    // ========== 持久化 ==========

    /// <summary>提交所有变更到数据库</summary>
    Task<int> SaveChangesAsync(CancellationToken ct = default);
}

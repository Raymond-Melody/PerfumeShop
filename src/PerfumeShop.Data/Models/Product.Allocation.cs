using System.ComponentModel.DataAnnotations.Schema;

namespace PerfumeShop.Data.Models;

/// <summary>
/// V20 扩展：费用分摊引擎所需的商品重量/体积属性
/// 对应 database/v20_expense_allocation_fix.sql 新增列
/// </summary>
public partial class Product
{
    /// <summary>商品重量（kg，含包装）</summary>
    [Column("Weight", TypeName = "decimal(9,3)")]
    public decimal? Weight { get; set; }

    /// <summary>商品包装体积（cm³）</summary>
    [Column("Volume", TypeName = "decimal(12,3)")]
    public decimal? Volume { get; set; }
}

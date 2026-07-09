using PerfumeShop.Data.Models;

namespace PerfumeShop.IntegrationTests;

/// <summary>
/// 数据模型验证测试 - 验证 EF Core 实体结构和基本逻辑
/// </summary>
public class ModelValidationTests
{
    [Fact]
    public void Product_DefaultValues_AreCorrect()
    {
        var product = new Product();
        Assert.Equal(0, product.ProductId);
        Assert.Null(product.IsActive); // bool? 默认为 null
        Assert.Null(product.ProductName);
    }

    [Fact]
    public void Order_TotalAmount_DefaultsToZero()
    {
        var order = new Order();
        Assert.Equal(0m, order.TotalAmount);
        Assert.Null(order.Status);
    }

    [Fact]
    public void OrderItem_CanBeCreated_WithRequiredFields()
    {
        var item = new OrderItem
        {
            OrderItemId = 1,
            OrderId = 100,
            ProductId = 50,
            Quantity = 2,
            UnitPrice = 199.99m
        };
        Assert.Equal(1, item.OrderItemId);
        Assert.Equal(100, item.OrderId);
        Assert.Equal(2, item.Quantity);
        Assert.Equal(199.99m, item.UnitPrice);
    }

    [Fact]
    public void FlashSale_DateRange_IsValid()
    {
        var now = DateTime.Now;
        var sale = new Data.Models.FlashSale
        {
            StartTime = now.AddHours(-1),
            EndTime = now.AddHours(1),
            IsActive = true
        };
        Assert.True(sale.StartTime <= now && sale.EndTime >= now);
    }

    [Fact]
    public void GroupBuyPlan_Properties_AreAccessible()
    {
        var plan = new GroupBuyPlan
        {
            PlanId = 1,
            ProductId = 10,
            GroupPrice = 99.9m,
            TeamSize = 3,
            DurationHours = 24,
            IsActive = true
        };
        Assert.Equal(99.9m, plan.GroupPrice);
        Assert.Equal(3, plan.TeamSize);
        Assert.True(plan.IsActive);
    }

    [Fact]
    public void SubscriptionPlan_Properties_AreAccessible()
    {
        var plan = new SubscriptionPlan
        {
            PlanId = 1,
            PlanName = "月度精选",
            Period = "月",
            Price = 299m,
            SampleCount = 3,
            FullSizeCount = 1,
            FreeShipping = true,
            IsActive = true
        };
        Assert.Equal("月度精选", plan.PlanName);
        Assert.Equal(299m, plan.Price);
        Assert.True(plan.FreeShipping);
    }

    [Fact]
    public void CommunityPost_DefaultCounts_AreZero()
    {
        var post = new CommunityPost
        {
            PostId = 1,
            UserId = 1,
            Title = "测试帖",
            Content = "测试内容",
            PostType = "Review",
            IsPublic = true,
            IsActive = true,
            CreatedAt = DateTime.Now
        };
        Assert.Equal(0, post.LikeCount);
        Assert.Equal(0, post.CommentCount);
        Assert.Equal(0, post.ViewCount);
        Assert.False(post.IsPinned);
    }

    [Fact]
    public void FragranceNote_HasNoteType()
    {
        var note = new FragranceNote
        {
            NoteId = 1,
            NoteName = "玫瑰",
            NoteType = "Middle",
            IsActive = true
        };
        Assert.Equal("Middle", note.NoteType);
        Assert.Equal("玫瑰", note.NoteName);
    }

    [Fact]
    public void CheckoutModel_ShippingFee_Calculation()
    {
        // 模拟运费计算逻辑: 满299免运费
        var subTotal1 = 200m;
        var fee1 = subTotal1 >= 299 ? 0 : 15;
        Assert.Equal(15, fee1);

        var subTotal2 = 300m;
        var fee2 = subTotal2 >= 299 ? 0 : 15;
        Assert.Equal(0, fee2);
    }

    [Fact]
    public void CartItem_PriceCalculation_IsCorrect()
    {
        var unitPrice = 199.99m;
        var quantity = 3;
        var total = unitPrice * quantity;
        Assert.Equal(599.97m, total);
    }

    [Fact]
    public void Pagination_Calculation_IsCorrect()
    {
        var totalCount = 100;
        var pageSize = 12;
        var totalPages = (int)Math.Ceiling((double)totalCount / pageSize);
        Assert.Equal(9, totalPages); // 100/12 = 8.33 → ceil = 9
    }

    [Fact]
    public void FlashSale_SoldPercentage_Calculation()
    {
        var stock = 20;
        var soldCount = 80;
        var soldPct = stock + soldCount > 0
            ? (int)((double)soldCount / (stock + soldCount) * 100) : 100;
        Assert.Equal(80, soldPct); // 80/(20+80) = 80%
    }
}

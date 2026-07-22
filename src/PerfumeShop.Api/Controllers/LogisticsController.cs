using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
public class LogisticsController : ControllerBase
{
    private readonly LogisticsRepository _repo;
    public LogisticsController(LogisticsRepository repo) => _repo = repo;

    [HttpGet("shipping")]
    public async Task<IActionResult> GetShippingOrders(int page = 1, int pageSize = 20, string? status = null)
    {
        var (items, total) = await _repo.GetShippingOrdersAsync(page, pageSize, status);
        return Ok(new { items, total });
    }

    [HttpGet("shipping/{id}")]
    public async Task<IActionResult> GetOrder(int id) { var o = await _repo.GetOrderAsync(id); return o == null ? NotFound() : Ok(o); }

    [HttpPut("shipping/{id}/status")]
    public async Task<IActionResult> UpdateStatus(int id, [FromBody] string status) { await _repo.UpdateOrderStatusAsync(id, status); return Ok(); }

    [HttpGet("returns")]
    public async Task<IActionResult> GetReturns(int page = 1, int pageSize = 20, string? status = null)
    {
        var (items, total) = await _repo.GetReturnsAsync(page, pageSize, status);
        return Ok(new { items, total });
    }

    [HttpGet("carriers")]
    public async Task<IActionResult> GetCarriers() => Ok(await _repo.GetCarriersAsync());

    [HttpPost("carriers")]
    public async Task<IActionResult> CreateCarrier([FromBody] ShippingCompany carrier) { await _repo.SaveCarrierAsync(carrier); return Ok(carrier); }

    [HttpPut("carriers/{id}")]
    public async Task<IActionResult> UpdateCarrier(int id, [FromBody] ShippingCompany carrier) { carrier.CompanyId = id; await _repo.SaveCarrierAsync(carrier); return Ok(); }

    [HttpDelete("carriers/{id}")]
    public async Task<IActionResult> DeleteCarrier(int id) { await _repo.DeleteCarrierAsync(id); return Ok(); }
}

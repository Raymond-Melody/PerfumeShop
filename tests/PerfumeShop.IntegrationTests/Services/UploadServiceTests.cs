using Microsoft.AspNetCore.Http;
using Moq;
using PerfumeShop.Shared.Services;

namespace PerfumeShop.IntegrationTests.Services;

/// <summary>UploadService 单元测试（正常路径 + 异常路径 + 边界条件）</summary>
public class UploadServiceTests
{
    private readonly Mock<IStorageProvider> _mockStorage;
    private readonly UploadService _service;

    public UploadServiceTests()
    {
        _mockStorage = new Mock<IStorageProvider>();
        _service = new UploadService(_mockStorage.Object, maxSizeBytes: 10 * 1024 * 1024);
    }

    // 辅助方法：构造 IFormFile Mock
    private static IFormFile CreateMockFile(byte[] content, string fileName, string contentType = "image/png")
    {
        var ms = new MemoryStream(content);
        return new FormFile(ms, 0, content.Length, "file", fileName)
        {
            Headers = new HeaderDictionary(),
            ContentType = contentType
        };
    }

    // --- PNG 魔数 (89 50 4E 47) ---
    private static byte[] ValidPngBytes()
    {
        var data = new byte[100];
        data[0] = 0x89; data[1] = 0x50; data[2] = 0x4E; data[3] = 0x47;
        return data;
    }

    // --- JPEG 魔数 (FF D8 FF) ---
    private static byte[] ValidJpegBytes()
    {
        var data = new byte[100];
        data[0] = 0xFF; data[1] = 0xD8; data[2] = 0xFF;
        return data;
    }

    [Fact]
    public async Task UploadAsync_ValidPngImage_ReturnsSuccess()
    {
        // Arrange
        var file = CreateMockFile(ValidPngBytes(), "photo.png");
        _mockStorage.Setup(s => s.SaveAsync(It.IsAny<byte[]>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>()))
                    .ReturnsAsync("/uploads/products/test.png");

        // Act
        var result = await _service.UploadAsync(file, "products", AllowedFileTypes.Images);

        // Assert
        Assert.True(result.Success);
        Assert.Equal("/uploads/products/test.png", result.Url);
        Assert.NotEmpty(result.FileName);
        Assert.Equal("image/png", result.ContentType);
        Assert.Equal(100, result.Size);
    }

    [Fact]
    public async Task UploadAsync_NullFile_ReturnsError()
    {
        // Act
        var result = await _service.UploadAsync(null!, "products");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("为空", result.Error);
    }

    [Fact]
    public async Task UploadAsync_EmptyFile_ReturnsError()
    {
        // Arrange
        var file = CreateMockFile(Array.Empty<byte>(), "empty.png");

        // Act
        var result = await _service.UploadAsync(file, "products");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("为空", result.Error);
    }

    [Fact]
    public async Task UploadAsync_FileExceedsMaxSize_ReturnsError()
    {
        // Arrange — 11MB 文件
        var largeData = new byte[11 * 1024 * 1024];
        var file = CreateMockFile(largeData, "large.png");

        // Act
        var result = await _service.UploadAsync(file, "products");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("超过限制", result.Error);
    }

    [Fact]
    public async Task UploadAsync_DisallowedExtension_ReturnsError()
    {
        // Arrange
        var file = CreateMockFile(new byte[] { 1, 2, 3 }, "malware.exe", "application/octet-stream");

        // Act
        var result = await _service.UploadAsync(file, "products", AllowedFileTypes.Images);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("不允许", result.Error);
    }

    [Fact]
    public async Task UploadAsync_MagicBytesMismatch_ReturnsError()
    {
        // Arrange — PNG 扩展名但内容不是 PNG
        var fakeData = new byte[100];
        fakeData[0] = 0x00; // 非 PNG 魔数
        var file = CreateMockFile(fakeData, "fake.png");

        // Act
        var result = await _service.UploadAsync(file, "products", AllowedFileTypes.Images);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("魔数验证失败", result.Error);
    }

    [Fact]
    public async Task UploadAsync_ValidJpeg_ReturnsSuccess()
    {
        // Arrange
        var file = CreateMockFile(ValidJpegBytes(), "photo.jpg", "image/jpeg");
        _mockStorage.Setup(s => s.SaveAsync(It.IsAny<byte[]>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>()))
                    .ReturnsAsync("/uploads/products/test.jpg");

        // Act
        var result = await _service.UploadAsync(file, "products", AllowedFileTypes.Images);

        // Assert
        Assert.True(result.Success);
        Assert.Equal("image/jpeg", result.ContentType);
    }

    [Fact]
    public async Task UploadAsync_AllowedDocumentType_ReturnsSuccess()
    {
        // Arrange
        var pdfData = new byte[] { 0x25, 0x50, 0x44, 0x46, 0x00 }; // %PDF
        var file = CreateMockFile(pdfData, "doc.pdf", "application/pdf");
        _mockStorage.Setup(s => s.SaveAsync(It.IsAny<byte[]>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>()))
                    .ReturnsAsync("/uploads/docs/doc.pdf");

        // Act
        var result = await _service.UploadAsync(file, "docs", AllowedFileTypes.Documents);

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void SanitizeFileName_RemovesDangerousCharacters()
    {
        // Act
        var result = UploadService.SanitizeFileName("../../etc/passwd");

        // Assert
        Assert.DoesNotContain("..", result);
        Assert.DoesNotContain("/", result);
        Assert.DoesNotContain("\\", result);
    }

    [Fact]
    public void GenerateUniqueFileName_ProducesCorrectFormat()
    {
        // Act
        var result = UploadService.GenerateUniqueFileName("test.png");

        // Assert
        Assert.EndsWith(".png", result);
        Assert.Contains("_", result);
        Assert.True(result.Length > 10);
    }

    [Fact]
    public void ValidateMagicBytes_ValidPng_ReturnsTrue()
    {
        Assert.True(UploadService.ValidateMagicBytes(ValidPngBytes(), ".png"));
    }

    [Fact]
    public void ValidateMagicBytes_InvalidData_ReturnsFalse()
    {
        Assert.False(UploadService.ValidateMagicBytes(new byte[] { 0, 0, 0, 0 }, ".png"));
    }
}

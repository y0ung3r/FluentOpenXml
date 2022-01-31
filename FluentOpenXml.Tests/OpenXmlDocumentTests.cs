using System.IO;
using FluentAssertions;
using Xunit;

namespace FluentOpenXml.Tests;

/// <summary>
/// Тесты для <see cref="OpenXmlDocument" />
/// </summary>
public class OpenXmlDocumentTests
{
    [Fact]
    public void Creating_a_new_document_with_stream()
    {
        // Arrange
        var stream = new MemoryStream();
        var sut = new OpenXmlDocument(stream);

        // Assert
        stream.Should().BeReadable();
        sut.Should().NotBeNull();

        sut.Close();
    }

    [Fact]
    public void Saving_a_document()
    {
        // Arrange
        // TODO: отрефакторить путь к файлу
        var filepath = @"D:\C#\FluentOpenXml\FluentOpenXml.Tests\Data\1.docx";
        var bytes = File.ReadAllBytes(filepath);
        var stream = new MemoryStream(bytes);
        var sut = new OpenXmlDocument();

        // Act
        sut.Save();

        // Assert
        stream.Length.Should().BeGreaterThan(0L);

        sut.Close();
    }

    [Fact]
    public void Closing_a_document()
    {
        // Arrange
        var stream = new MemoryStream();
        var sut = new OpenXmlDocument(stream);

        // Act
        sut.Close();

        // Assert
        stream.Should().NotBeReadable();
    }

    [Fact]
    public void Loading_a_document_from_path()
    {
        // Arrange
        // TODO: отрефакторить путь к файлу
        var filepath = @"D:\C#\FluentOpenXml\FluentOpenXml.Tests\Data\1.docx";
        var stream = new FileStream(filepath, FileMode.Open, FileAccess.ReadWrite);
        var sut = new OpenXmlDocument();

        // Act
        sut.LoadFrom(stream);

        // Assert
        sut.Should().NotBeNull();
        stream.Length.Should().BeGreaterThan(0L);

        sut.Close();
    }
}
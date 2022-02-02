using System;
using System.IO;
using FluentAssertions;
using FluentOpenXml.Exceptions;
using Xunit;

namespace FluentOpenXml.Tests;

/// <summary>
/// Тесты инициализации для <see cref="OpenXmlDocument" />
/// </summary>
public class CreationTests
{
    #region Создание новых документов

    [Fact]
    public void Creating_with_default_settings()
    {
        // Arrange
        var sut = new OpenXmlDocument();
        
        // Assert
        sut.Settings.Should().BeEquivalentTo(DocumentSettings.Default);
        sut.IsEmpty.Should().BeTrue();
        sut.Should().NotBeNull();
    }

    [Fact]
    public void Creating_with_settings()
    {
        // Arrange
        var settings = new DocumentSettings()
        {
            IsReadOnly = false,
            AllowAutoSaving = true,
            AllowUpdateFieldsOnOpen = true
        };

        // Act
        var sut = new OpenXmlDocument(settings);

        // Assert
        sut.Settings.Should().BeEquivalentTo(settings);
        sut.IsEmpty.Should().BeTrue();
        sut.Should().NotBeNull();
    }

    [Fact]
    public void Loading_with_default_settings()
    {
        // Arrange
        var filepath = @"D:\C#\FluentOpenXml\FluentOpenXml.Tests\Data\1.docx";
        
        // Act
        var sut = new OpenXmlDocument(filepath);
        
        // Assert
        sut.Settings.Should().BeEquivalentTo(DocumentSettings.Default);
        sut.IsEmpty.Should().BeFalse();
    }
    
    [Fact]
    public void Loading_with_settings()
    {
        // Arrange
        var filepath = @"D:\C#\FluentOpenXml\FluentOpenXml.Tests\Data\1.docx";
        var settings = new DocumentSettings()
        {
            IsReadOnly = false,
            AllowAutoSaving = true,
            AllowUpdateFieldsOnOpen = true
        };
        
        // Act
        var sut = new OpenXmlDocument(filepath, settings);
        
        // Assert
        sut.Settings.Should().BeEquivalentTo(settings);
        sut.IsEmpty.Should().BeFalse();
    }

    [Fact]
    public void Obtaining_exceptions_for_invalid_data()
    {
        this.Invoking
        (
            _ => new OpenXmlDocument
            (
                default(DocumentSettings)
            )
        )
        .Should()
        .Throw<ArgumentNullException>();

        this.Invoking
        (
            _ => new OpenXmlDocument(string.Empty)
        )
        .Should()
        .Throw<ArgumentException>();
        
        this.Invoking
        (
            _ => new OpenXmlDocument
            (
                string.Empty, 
                DocumentSettings.Default
            )
        )
        .Should()
        .Throw<ArgumentException>();

        this.Invoking
        (
            _ => new OpenXmlDocument(Stream.Null)
        )
        .Should()
        .Throw<InvalidDocumentException>();
        
        this.Invoking
        (
            _ => new OpenXmlDocument
            (
                Stream.Null, 
                DocumentSettings.Default
            )
        )
        .Should()
        .Throw<InvalidDocumentException>();
    }
    
    #endregion
}
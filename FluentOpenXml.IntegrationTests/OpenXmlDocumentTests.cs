using System;
using System.IO;
using FluentAssertions;
using FluentOpenXml.Exceptions;
using Xunit;

namespace FluentOpenXml.IntegrationTests;

/// <summary>
/// Интеграционные тесты для <see cref="OpenXmlDocument"/>
/// </summary>
public class OpenXmlDocumentTests
{
	[Fact] 
	public void Should_set_default_settings()
	{
		// Arrange & Act
		var sut = new OpenXmlDocument();

		// Assert
		sut.Settings.Should().BeEquivalentTo(DocumentSettings.Default);
	}
	
	[Fact] 
	public void Should_set_custom_settings()
	{
		// Arrange
		var settings = new DocumentSettings()
		{
			IsReadOnly = true
		};
		
		// Act
		var sut = new OpenXmlDocument(settings);

		// Assert
		sut.Settings.Should().NotBeEquivalentTo(DocumentSettings.Default);
	}
	
	[Fact]
	public void Should_be_empty_when_creating_a_new_document()
	{
		// Arrange & Act
		var sut = new OpenXmlDocument();

		// Assert
		sut.IsEmpty.Should().Be(true);
	}

	[Fact]
	public void Should_not_be_empty_when_opening_a_document()
	{
		// Arrange & Act
		var filepath = Path.Combine(AppContext.BaseDirectory, "Data", "Should_not_be_empty_when_opening_a_document.docx");
		var sut = new OpenXmlDocument(filepath);
		
		// Assert
		sut.IsEmpty.Should().Be(false);
	}

	[Fact]
	public void Should_not_be_saved_when_opened_in_readonly_mode()
	{
		// Arrange
		var settings = new DocumentSettings()
		{
			IsReadOnly = true
		};
		
		// Act
		var sut = new OpenXmlDocument(settings);
		
		// Assert
		sut.Invoking
		(
			x => x.Save()
		)
		.Should()
		.Throw<DocumentInReadOnlyModeException>();
	}

	[Fact]
	public void Should_not_be_editable_when_opened_in_readonly_mode()
	{
		// Arrange
		var settings = new DocumentSettings()
		{
			IsReadOnly = true
		};
		
		// Act
		var sut = new OpenXmlDocument(settings);
		
		// Assert
		sut.Invoking
		(
			x => x.Edit(_ => { })
		)
		.Should()
		.Throw<DocumentInReadOnlyModeException>();
	}

	[Fact]
	public void Should_create_DocumentBuilder_when_document_editing()
	{
		// Arrange & Act
		var sut = new OpenXmlDocument();
		
		// Assert
		sut.Edit
		(
			x => x.Should().NotBeNull()
		);
	}

	[Fact]
	public void Should_stay_open_after_save()
	{
		// Arrange
		var filepath = Path.Combine(AppContext.BaseDirectory, "Data", "Should_stay_open_after_save.docx");
		var sut = new OpenXmlDocument(filepath);

		// Act
		sut.Save();

		// Assert
		sut.Invoking
		(
			x => x.Edit(_ => { })
		)
		.Should()
		.NotThrow<ObjectDisposedException>();
	}

	[Fact]
	public void Should_save_to_path()
	{
		// Arrange
		var filepath = Path.Combine(AppContext.BaseDirectory, "Data", "Should_save_to_path.docx");
		var sut = new OpenXmlDocument();
		File.Delete(filepath);

		// Act
		sut.SaveTo(filepath);
		
		// Assert
		File.Exists(filepath).Should().Be(true);
	}

	[Fact]
	public void Should_save_to_stream()
	{
		// Arrange
		var filepath = Path.Combine(AppContext.BaseDirectory, "Data", "Should_save_to_stream.docx");
		var sut = new OpenXmlDocument(filepath);
		var destinationStream = new MemoryStream();
		
		// Act
		sut.SaveTo(destinationStream);

		// Assert
		destinationStream.Length.Should().BePositive();
	}

	[Fact]
	public void Should_unload_document_after_close()
	{
		// Arrange
		var sut = new OpenXmlDocument();
		
		// Act
		sut.Close();
		
		// Assert
		sut.Settings.Should().BeNull();
		
		sut.Invoking(x => x.IsEmpty)
		   .Should()
		   .Throw<ObjectDisposedException>();
		
		sut.Invoking
		(
			x => x.Edit(_ => { })
		)
		.Should()
		.Throw<ObjectDisposedException>();
	}
}
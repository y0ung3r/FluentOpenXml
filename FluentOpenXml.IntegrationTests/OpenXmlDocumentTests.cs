using System;
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
	public void Should_empty_when_creating_a_new_document()
	{
		// Arrange & Act
		var sut = new OpenXmlDocument();

		// Assert
		sut.IsEmpty.Should().Be(true);
	}

	[Fact]
	public void Should_unload_after_closing()
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
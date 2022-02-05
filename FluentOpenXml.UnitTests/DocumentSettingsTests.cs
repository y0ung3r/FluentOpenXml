using System.IO;
using FluentAssertions;
using Xunit;

namespace FluentOpenXml.UnitTests;

/// <summary>
/// Тесты для <see cref="DocumentSettings"/>
/// </summary>
public class DocumentSettingsTests
{
	[Fact]
	public void Different_settings_must_have_difference_references()
	{
		// Arrange & Act
		var sut = DocumentSettings.Default;
		var anotherSut = DocumentSettings.Default;

		// Assert
		sut.Should().NotBeSameAs(anotherSut);
	}
	
	[Fact]
	public void Should_activate_readonly_mode_if_IsReadOnly_setting_is_true()
	{
		// Arrange & Act
		var sut = new DocumentSettings()
		{
			IsReadOnly = true
		};
		
		// Assert
		sut.DocumentMode.Should().Be(FileMode.Open);
		sut.DocumentAccess.Should().Be(FileAccess.Read);
	}
	
	[Fact]
	public void Should_activate_write_mode_if_IsReadOnly_setting_is_false()
	{
		// Arrange & Act
		var sut = new DocumentSettings()
		{
			IsReadOnly = false
		};
		
		// Assert
		sut.DocumentMode.Should().Be(FileMode.OpenOrCreate);
		sut.DocumentAccess.Should().Be(FileAccess.ReadWrite);
	}

	[Fact]
	public void Should_disallow_auto_saving_if_IsReadOnly_setting_is_true()
	{
		// Arrange & Act
		var sut = new DocumentSettings()
		{
			IsReadOnly = true,
			AllowAutoSaving = true
		};

		// Assert
		sut.AllowAutoSaving.Should().Be(false);
	}
	
	[Fact]
	public void Should_not_affect_to_auto_saving_if_IsReadOnly_setting_is_false()
	{
		// Arrange & Act
		var sut = new DocumentSettings()
		{
			IsReadOnly = false,
			AllowAutoSaving = true
		};

		// Assert
		sut.AllowAutoSaving.Should().Be(true);
	}
}
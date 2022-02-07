using FluentAssertions;
using FluentOpenXml.Units;
using Xunit;

namespace FluentOpenXml.UnitTests.Units;

/// <summary>
/// Тесты для <see cref="Centimeters"/>
/// </summary>
public class CentimetersTests
{
	[Fact]
	public void Should_convert_to_emu()
	{
		// Arrange
		var sut = new Centimeters(10.0);

		// Act
		var emu = sut.ToEmu();
		
		// Assert
		emu.Value.Should().Be(3600000.0);
	}
	
	[Fact]
	public void Should_convert_to_millimeters()
	{
		// Arrange
		var sut = new Centimeters(10.0);

		// Act
		var millimeters = sut.ToMillimeters();
		
		// Assert
		millimeters.Value.Should().Be(100.0);
	}
	
	[Fact]
	public void Should_convert_to_points()
	{
		// Arrange
		var sut = new Centimeters(10.0);

		// Act
		var points = sut.ToPoints();
		
		// Assert
		points.Value.Should().Be(283.46456692913387);
	}
}
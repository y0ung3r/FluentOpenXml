using FluentAssertions;
using FluentOpenXml.Units;
using Xunit;

namespace FluentOpenXml.UnitTests.Units;

/// <summary>
/// Тесты для <see cref="Dxa"/>
/// </summary>
public class DxaTests
{
	[Fact]
	public void Should_convert_to_emu()
	{
		// Arrange
		var sut = new Dxa(10.0);

		// Act
		var emu = sut.ToEmu();
		
		// Assert
		emu.Value.Should().Be(6350.0);
	}
	
	[Fact]
	public void Should_convert_to_points()
	{
		// Arrange
		var sut = new Dxa(10.0);

		// Act
		var points = sut.ToPoints();
		
		// Assert
		points.Value.Should().Be(0.5);
	}
	
	[Fact]
	public void Should_convert_to_inches()
	{
		// Arrange
		var sut = new Dxa(10.0);

		// Act
		var inches = sut.ToInches();
		
		// Assert
		inches.Value.Should().Be(0.1388888888888889);
	}
}
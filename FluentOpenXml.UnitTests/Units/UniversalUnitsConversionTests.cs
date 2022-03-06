using FluentAssertions;
using FluentOpenXml.Units.Universal;
using Xunit;

namespace FluentOpenXml.UnitTests.Units;

/// <summary>
/// Тесты для <see cref="UniversalUnits" />
/// </summary>
public class UniversalUnitsConversionTests
{
	[Fact]
	public void Should_twips_convert_to_emu()
	{
		// Arrange
		var sut = new Twips(567.0);

		// Act
		var twips = sut.ToEmu();

		// Assert
		twips.Value.Should().Be(360045.0);
	}
	
	[Fact]
	public void Should_centimeters_convert_to_emu()
	{
		// Arrange
		var sut = new Centimeters(10.0);

		// Act
		var emu = sut.ToEmu();

		// Assert
		emu.Value.Should().Be(3600000.0);
	}
	
	[Fact]
	public void Should_points_convert_to_emu()
	{
		// Arrange
		var sut = new Points(10.0);

		// Act
		var emu = sut.ToEmu();

		// Assert
		emu.Value.Should().Be(127000.0);
	}
	
	[Fact]
	public void Should_points_convert_to_half_points()
	{
		// Arrange
		var sut = new Points(10.0);

		// Act
		var halfPoints = sut.ToHalfPoints();

		// Assert
		halfPoints.Value.Should().Be(20.0);
	}

	[Fact]
	public void Should_inches_convert_to_emu()
	{
		// Arrange
		var sut = new Inches(10.0);

		// Act
		var emu = sut.ToEmu();

		// Assert
		emu.Value.Should().Be(9144000.0);
	}

	[Fact]
	public void Should_millimeters_convert_to_emu()
	{
		// Arrange
		var sut = new Millimeters(10.0);

		// Act
		var emu = sut.ToEmu();

		// Assert
		emu.Value.Should().Be(360000.0);
	}
	
	[Fact]
	public void Should_picas_convert_to_emu()
	{
		// Arrange
		var sut = new Picas(10.0);

		// Act
		var emu = sut.ToEmu();

		// Assert
		emu.Value.Should().Be(1524000.0);
	}
	
	[Fact]
	public void Should_picas_convert_to_inches()
	{
		// Arrange
		var sut = new Picas(15.0);

		// Act
		var inches = sut.ToInches();

		// Assert
		inches.Value.Should().Be(2.5);
	}
}
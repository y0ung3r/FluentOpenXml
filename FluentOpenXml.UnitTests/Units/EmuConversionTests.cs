using System;
using FluentAssertions;
using FluentOpenXml.Units;
using FluentOpenXml.Units.Universal;
using Xunit;

namespace FluentOpenXml.UnitTests.Units;

/// <summary>
/// Тесты для <see cref="Emu"/>
/// </summary>
public class EmuConversionTests
{
	[Fact]
	public void Should_emu_convert_to_centimeters()
	{
		// Arrange
		var sut = new Emu(360000.0);

		// Act
		var centimeters = sut.ToCentimeters();

		// Assert
		centimeters.Value.Should().Be(1.0);
	}
	
	[Fact]
	public void Should_emu_convert_to_points()
	{
		// Arrange
		var sut = new Emu(101600.0);

		// Act
		var points = sut.ToPoints();

		// Assert
		points.Value.Should().Be(8.0);
	}

	[Fact]
	public void Should_emu_convert_to_inches()
	{
		// Arrange
		var sut = new Emu(914400.0);

		// Act
		var inches = sut.ToInches();

		// Assert
		inches.Value.Should().Be(1.0);
	}
	
	[Fact]
	public void Should_emu_convert_to_picas()
	{
		// Arrange
		var sut = new Emu(152400.0);

		// Act
		var picas = sut.ToPicas();

		// Assert
		picas.Value.Should().Be(1.0);
	}
	
	[Fact]
	public void Should_emu_convert_to_millimeters()
	{
		// Arrange
		var sut = new Emu(360000.0);

		// Act
		var millimeters = sut.ToMillimeters();

		// Assert
		millimeters.Value.Should().Be(10.0);
	}
	
	[Fact]
	public void Should_emu_convert_to_twips()
	{
		// Arrange
		var sut = new Emu(360045.0);

		// Act
		var twips = sut.ToTwips();

		// Assert
		twips.Value.Should().Be(567.0);
	}

	[Fact]
	public void Should_emu_convert_to_centimeters_with_reflection()
	{
		// Arrange
		var sut = new Emu(360000.0);

		// Act
		var centimeters = sut.To<Centimeters>();

		// Assert
		centimeters.Value.Should().Be(1.0);
	}
	
	[Fact]
	public void Should_emu_convert_to_points_with_reflection()
	{
		// Arrange
		var sut = new Emu(101600.0);

		// Act
		var points = sut.To<Points>();

		// Assert
		points.Value.Should().Be(8.0);
	}
	
	[Fact]
	public void Should_emu_convert_to_inches_with_reflection()
	{
		// Arrange
		var sut = new Emu(914400.0);

		// Act
		var inches = sut.To<Inches>();

		// Assert
		inches.Value.Should().Be(1.0);
	}
	
	[Fact]
	public void Should_emu_convert_to_picas_with_reflection()
	{
		// Arrange
		var sut = new Emu(152400.0);

		// Act
		var picas = sut.ToPicas();

		// Assert
		picas.Value.Should().Be(1.0);
	}
	
	[Fact]
	public void Should_emu_convert_to_millimeters_with_reflection()
	{
		// Arrange
		var sut = new Emu(360000.0);

		// Act
		var millimeters = sut.To<Millimeters>();

		// Assert
		millimeters.Value.Should().Be(10.0);
	}
	
	[Fact]
	public void Should_emu_convert_to_twips_with_reflection()
	{
		// Arrange
		var sut = new Emu(360045.0);

		// Act
		var twips = sut.To<Twips>();
		
		// Assert
		twips.Value.Should().Be(567.0);
	}

	[Fact]
	public void Should_throw_exception_if_conversion_method_not_defined()
	{
		// Arrange
		var sut = new Emu(360000.0);

		// Act && Assert
		sut.Invoking
		(
			emu => emu.To<UniversalUnits>()
		)
		.Should()
		.Throw<ArgumentException>();
	}
}
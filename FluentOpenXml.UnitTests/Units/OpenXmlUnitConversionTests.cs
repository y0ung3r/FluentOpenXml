﻿using FluentAssertions;
using FluentOpenXml.Units.OpenXml;
using Xunit;

namespace FluentOpenXml.UnitTests.Units;

/// <summary>
/// Тесты для <see cref="OpenXmlUnits"/>
/// </summary>
public class OpenXmlUnitConversionTests
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
		var sut = new Emu(360000.0);

		// Act
		var twips = sut.ToTwips();

		// Assert
		twips.Value.Should().Be(567L);
	}
	
	[Fact]
	public void Should_twips_convert_to_emu()
	{
		// Arrange
		var sut = new Twips(567.0);

		// Act
		var twips = sut.ToEmu();

		// Assert
		twips.Value.Should().Be(360045L);
	}
}
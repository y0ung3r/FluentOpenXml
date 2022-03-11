using System;
using FluentAssertions;
using FluentOpenXml.Units;
using FluentOpenXml.Units.Extensions;
using FluentOpenXml.Units.Universal;
using Xunit;

namespace FluentOpenXml.UnitTests.Units;

/// <summary>
/// Тесты для <see cref="DoubleExtensions"/>
/// </summary>
public class DoubleExtensionsTests
{
	[Fact]
	public void Should_present_double_value_as_centimeters()
	{
		// Arrange
		var sut = 15.0;

		// Act
		var centimeters = sut.As<Centimeters>();

		// Assert
		centimeters.Value.Should().Be(15.0);
	}
	
	[Fact]
	public void Should_present_double_value_as_points()
	{
		// Arrange
		var sut = 15.0;

		// Act
		var points = sut.As<Points>();

		// Assert
		points.Value.Should().Be(15.0);
	}
	
	[Fact]
	public void Should_present_double_value_as_inches()
	{
		// Arrange
		var sut = 15.0;

		// Act
		var inches = sut.As<Inches>();

		// Assert
		inches.Value.Should().Be(15.0);
	}
	
	[Fact]
	public void Should_present_double_value_as_picas()
	{
		// Arrange
		var sut = 15.0;

		// Act
		var picas = sut.As<Picas>();

		// Assert
		picas.Value.Should().Be(15.0);
	}
	
	[Fact]
	public void Should_present_double_value_as_millimeters()
	{
		// Arrange
		var sut = 15.0;

		// Act
		var millimeters = sut.As<Millimeters>();

		// Assert
		millimeters.Value.Should().Be(15.0);
	}
	
	[Fact]
	public void Should_throw_exception_if_trying_pass_abstract_unit()
	{
		// Arrange
		var sut = 15.0;

		// Act && Assert
		sut.Invoking
		(
			x => x.As<FloatingPointUnits>()
		)
		.Should()
		.Throw<ArgumentException>();
	}
}
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
	public void Should_present_double_value_as_any_universal_units()
	{
		// Arrange
		var sut = 15.0;

		// Act
		var anyUnits = sut.As<Centimeters>();

		// Assert
		anyUnits.Value.Should().Be(15.0);
	}

	[Fact]
	public void Should_present_double_value_as_any_universal_units_and_convert_it_to_twips()
	{
		// Arrange
		var sut = 1.0;

		// Act
		var anyUnits = sut.ToTwipsAs<Inches>();

		// Assert
		anyUnits.Value.Should().Be(1440.0);
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
using FluentAssertions;
using FluentOpenXml.Units.Relative;
using Xunit;

namespace FluentOpenXml.UnitTests.Units;

/// <summary>
/// Тесты для относительных единиц измерения
/// </summary>
public class RelativeUnitsConversionTests
{
	[Fact]
	public void Should_fiftieths_of_a_percentage_convert_to_percentages()
	{
		// Arrange
		var sut = new FiftiethsOfAPercentage(5000.0);

		// Act
		var percentages = sut.ToPercentages();

		// Assert
		percentages.Value.Should().Be(100.0);
	}
	
	[Fact]
	public void Should_percentages_convert_to_fiftieths_of_a_percentage()
	{
		// Arrange
		var sut = new Percentages(50.0);

		// Act
		var fiftiethsOfAPercentage = sut.ToFiftiethsOfAPercentage();

		// Assert
		fiftiethsOfAPercentage.Value.Should().Be(2500.0);
	}
}
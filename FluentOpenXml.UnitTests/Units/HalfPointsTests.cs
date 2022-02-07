using FluentAssertions;
using FluentOpenXml.Units;
using Xunit;

namespace FluentOpenXml.UnitTests.Units;

/// <summary>
/// Тесты для <see cref="HalfPoints"/>
/// </summary>
public class HalfPointsTests
{
	[Fact]
	public void Should_convert_to_points()
	{
		// Arrange
		var sut = new HalfPoints(10.0);

		// Act
		var points = sut.ToPoints();
		
		// Assert
		points.Value.Should().Be(5.0);
	}
}
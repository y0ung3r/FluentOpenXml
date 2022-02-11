using FluentAssertions;
using FluentOpenXml.Units;
using Xunit;

namespace FluentOpenXml.UnitTests.Units;

/// <summary>
/// Тесты для <see cref="Points"/>
/// </summary>
public class PointsTests
{
	[Fact]
	public void Should_convert_to_emu()
	{
		// Arrange
		var sut = new Points(10.0);

		// Act
		var emu = sut.ToEmu();
		
		// Assert
		emu.Value.Should().Be(127000L);
	}
	
	[Fact]
	public void Should_convert_to_dxa()
	{
		// Arrange
		var sut = new Points(10.0);

		// Act
		var dxa = sut.ToDxa();
		
		// Assert
		dxa.Value.Should().Be(200.0);
	}
}
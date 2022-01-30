using Xunit;

namespace FluentOpenXml.Tests;

/// <summary>
/// Тесты для <see cref="OpenXmlDocument" />
/// </summary>
public class OpenXmlDocumentTests
{
	[Fact]
	public void Creating_test()
	{
		var sut = new OpenXmlDocument();
		
		Assert.NotNull(sut);
	}
}
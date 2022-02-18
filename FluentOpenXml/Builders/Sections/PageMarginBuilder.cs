using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Sections.Interfaces;
using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.Sections;

/// <summary>
/// Стандартная реализация <see cref="IPageMarginBuilder"/>
/// </summary>
internal class PageMarginBuilder : OpenXmlElementBuilder, IPageMarginBuilder
{
	/// <summary>
	/// Размеры страницы
	/// </summary>
	private readonly PageMargin _pageMargin;

	/// <summary>
	/// Инициализирует <see cref="PageSizeBuilder"/> по указанному <see cref="MainDocumentPart"/> и <see cref="PageMargin"/>
	/// </summary>
	/// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
	/// <param name="pageMargin">Внешние отступы страницы</param>
	public PageMarginBuilder(MainDocumentPart mainDocumentPart, PageMargin pageMargin) 
		: base(mainDocumentPart)
	{
		_pageMargin = pageMargin;
	}
	
	/// <inheritdoc />
	public IPageMarginBuilder SetLeft(UniversalUnits value)
	{
		_pageMargin.Left = value.ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetTop(UniversalUnits value)
	{
		_pageMargin.Top = value.ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetRight(UniversalUnits value)
	{
		_pageMargin.Right = value.ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetBottom(UniversalUnits value)
	{
		_pageMargin.Bottom = value.ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetHeader(UniversalUnits value)
	{
		_pageMargin.Header = value.ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetFooter(UniversalUnits value)
	{
		_pageMargin.Footer = value.ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetGutter(UniversalUnits value)
	{
		_pageMargin.Gutter = value.ToEmu().ToTwips();

		return this;
	}
}
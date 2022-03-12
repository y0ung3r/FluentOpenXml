using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Sections.Enums;
using FluentOpenXml.Builders.Sections.PageLayout.Interfaces;

namespace FluentOpenXml.Builders.Sections.PageLayout;

/// <summary>
/// Стандартная реализация <see cref="IPageLayoutBuilder"/>
/// </summary>
internal class PageLayoutBuilder : OpenXmlElementBuilder, IPageLayoutBuilder
{
	/// <summary>
	/// Размеры страницы
	/// </summary>
	private readonly PageSize _size;

	/// <summary>
	/// Отступы страницы
	/// </summary>
	private readonly PageMargin _margin;

	/// <summary>
	/// Инициализирует <see cref="OpenXmlElementBuilder"/> по указанному <see cref="OpenXmlElementBuilder.MainDocumentPart"/>
	/// </summary>
	/// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
	/// <param name="size">Размеры страницы</param>
	/// <param name="margin">Отступы страницы</param>
	public PageLayoutBuilder(MainDocumentPart mainDocumentPart, PageSize size, PageMargin margin) 
		: base(mainDocumentPart)
	{
		_size = size;
		_margin = margin;
	}
	
	/// <inheritdoc />
	public IPageLayoutBuilder SetOrientation(PageOrientation orientation)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc />
	public IPageLayoutBuilder SetSize(Action<IPageSizeBuilder> setSize)
	{
		ConfigureWith<PageSizeBuilder>(setSize, _size);
		
		return this;
	}

	/// <inheritdoc />
	public IPageLayoutBuilder SetMargin(Action<IPageMarginBuilder> setMargin)
	{
		ConfigureWith<PageMarginBuilder>(setMargin, _margin);
		
		return this;
	}
}
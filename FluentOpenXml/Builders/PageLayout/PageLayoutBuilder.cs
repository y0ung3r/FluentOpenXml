using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.PageLayout.Interfaces;
using FluentOpenXml.Builders.Sections.Enums;

namespace FluentOpenXml.Builders.PageLayout;

/// <summary>
/// Стандартная реализация <see cref="IPageLayoutBuilder"/>
/// </summary>
internal class PageLayoutBuilder : OpenXmlElementBuilder, IPageLayoutBuilder
{
	/// <summary>
	/// Размеры страницы
	/// </summary>
	private readonly PageSize _pageSize;

	/// <summary>
	/// Отступы страницы
	/// </summary>
	private readonly PageMargin _pageMargin;

	/// <summary>
	/// Ориентация страницы
	/// </summary>
	private PageOrientation Orientation
	{
		get
		{
			_pageSize.Orient ??= PageOrientationValues.Portrait;
			return (PageOrientation)_pageSize.Orient.Value;
		}
        
		set => _pageSize.Orient = (PageOrientationValues)value;
	}
	
	/// <summary>
	/// Инициализирует <see cref="OpenXmlElementBuilder"/> по указанному <see cref="OpenXmlElementBuilder.MainDocumentPart"/>
	/// </summary>
	/// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
	/// <param name="pageSize">Размеры страницы</param>
	/// <param name="pageMargin">Отступы страницы</param>
	public PageLayoutBuilder(MainDocumentPart mainDocumentPart, PageSize pageSize, PageMargin pageMargin) 
		: base(mainDocumentPart)
	{
		_pageSize = pageSize;
		_pageMargin = pageMargin;
	}

	/// <inheritdoc />
	public IPageLayoutBuilder SetOrientation(PageOrientation orientation)
	{
		if (Orientation != orientation)
		{
			Transpose();
		}

		return this;
	}

	/// <inheritdoc />
	public IPageLayoutBuilder SetSize(Action<IPageSizeBuilder> setSize)
	{
		ArgumentNullException.ThrowIfNull(setSize);
		ConfigureWith<PageSizeBuilder>(setSize, _pageSize);
		
		return this;
	}

	/// <inheritdoc />
	public IPageLayoutBuilder SetMargin(Action<IPageMarginBuilder> setMargin)
	{
		ArgumentNullException.ThrowIfNull(setMargin);
		ConfigureWith<PageMarginBuilder>(setMargin, _pageMargin);
		
		return this;
	}

	/// <summary>
	/// Сначала переставляет местами значения размеров страницы, а затем переставляет значения отступов
	/// </summary>
	public void Transpose()
	{
		var transpose = new Action<ITransposable>
		(
			transposable => transposable.Transpose()
		);

		SetSize(transpose);
		SetMargin(transpose);
	}
}
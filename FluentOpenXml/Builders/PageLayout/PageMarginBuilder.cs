using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.PageLayout.Interfaces;
using FluentOpenXml.Units.Extensions;
using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.PageLayout;

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
	/// Левый отступ
	/// </summary>
	private Twips Left
	{
		get => new Twips(_pageMargin.Left);
		set => _pageMargin.Left = value;
	}
	
	/// <summary>
	/// Верхний отступ
	/// </summary>
	private Twips Top
	{
		get => new Twips(_pageMargin.Top);
		set => _pageMargin.Top = value;
	}

	/// <summary>
	/// Правый отступ
	/// </summary>
	private Twips Right
	{
		get => new Twips(_pageMargin.Right);
		set => _pageMargin.Right = value;
	}
	
	/// <summary>
	/// Нижний отступ
	/// </summary>
	private Twips Bottom
	{
		get => new Twips(_pageMargin.Bottom);
		set => _pageMargin.Bottom = value;
	}
	
	/// <summary>
	/// Отступ от верхнего колонтитула
	/// </summary>
	private Twips Header
	{
		get => new Twips(_pageMargin.Header);
		set => _pageMargin.Header = value;
	}
	
	/// <summary>
	/// Отступ от нижнего колонтитула
	/// </summary>
	private Twips Footer
	{
		get => new Twips(_pageMargin.Footer);
		set => _pageMargin.Footer = value;
	}
	
	/// <summary>
	/// Интервал переплета страниц
	/// </summary>
	private Twips Gutter
	{
		get => new Twips(_pageMargin.Gutter);
		set => _pageMargin.Gutter = value;
	}

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
	public IPageMarginBuilder SetLeft<TUnits>(double value)
		where TUnits : UniversalUnits
	{
		Left = value.As<TUnits>().ToEmu().ToTwips();
		
		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetTop<TUnits>(double value)
		where TUnits : UniversalUnits
	{
		Top = value.As<TUnits>().ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetRight<TUnits>(double value)
		where TUnits : UniversalUnits 
	{
		Right = value.As<TUnits>().ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetBottom<TUnits>(double value)
		where TUnits : UniversalUnits
	{
		Bottom = value.As<TUnits>().ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetHeader<TUnits>(double value)
		where TUnits : UniversalUnits 
	{
		Header = value.As<TUnits>().ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetFooter<TUnits>(double value)
		where TUnits : UniversalUnits 
	{
		Footer = value.As<TUnits>().ToEmu().ToTwips();

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetGutter<TUnits>(double value)
		where TUnits : UniversalUnits
	{
		Gutter = value.As<TUnits>().ToEmu().ToTwips();

		return this;
	}

	/// <summary>
	/// Меняет значение верхнего, нижнего, левого и правого отступов местами с соответственно левым, правым, нижним и верхним отступами
	/// </summary>
	public void Transpose()
	{
		var bottom = Math.Max(0.0, Bottom.Value).As<Twips>();
		var top = Math.Max(0.0, Top.Value).As<Twips>();
		(Top, Bottom, Left, Right) = (Left, Right, bottom, top);
	}
}
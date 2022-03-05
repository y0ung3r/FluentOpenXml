using System.Linq.Expressions;
using DocumentFormat.OpenXml;
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

	/// <summary>
	/// Устанавливает отступ для указанной стороны элемент
	/// </summary>
	/// <param name="expression">Путь к свойству</param>
	/// <param name="value">Значение</param>
	/// <typeparam name="TProperty">Тип свойства</typeparam>
	private IPageMarginBuilder SetMargin<TProperty>(Expression<Func<PageMargin, TProperty>> expression, UniversalUnits value)
	{
		SetPropertyValue<PageMargin, TProperty, UniversalUnits>
		(
			_pageMargin, 
			expression, 
			value, 
			units => units.ToEmu().ToTwips()
		);

		return this;
	}

	/// <inheritdoc />
	public IPageMarginBuilder SetLeft<TUnits>(TUnits value)
		where TUnits : UniversalUnits => SetMargin(margin => margin.Left, value);

	/// <inheritdoc />
	public IPageMarginBuilder SetTop<TUnits>(TUnits value)
		where TUnits : UniversalUnits => SetMargin(margin => margin.Top, value);

	/// <inheritdoc />
	public IPageMarginBuilder SetRight<TUnits>(TUnits value)
		where TUnits : UniversalUnits => SetMargin(margin => margin.Right, value);

	/// <inheritdoc />
	public IPageMarginBuilder SetBottom<TUnits>(TUnits value)
		where TUnits : UniversalUnits => SetMargin(margin => margin.Bottom, value);

	/// <inheritdoc />
	public IPageMarginBuilder SetHeader<TUnits>(TUnits value)
		where TUnits : UniversalUnits => SetMargin(margin => margin.Header, value);

	/// <inheritdoc />
	public IPageMarginBuilder SetFooter<TUnits>(TUnits value)
		where TUnits : UniversalUnits => SetMargin(margin => margin.Footer, value);

	/// <inheritdoc />
	public IPageMarginBuilder SetGutter<TUnits>(TUnits value)
		where TUnits : UniversalUnits => SetMargin(margin => margin.Gutter, value);
}
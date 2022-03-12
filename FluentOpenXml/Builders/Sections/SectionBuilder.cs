using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Paragraphs.Interfaces;
using FluentOpenXml.Builders.Sections.Interfaces;
using FluentOpenXml.Builders.Sections.PageLayout;
using FluentOpenXml.Builders.Sections.PageLayout.Interfaces;
using FluentOpenXml.Extensions;

namespace FluentOpenXml.Builders.Sections;

/// <summary>
/// Стандартная реализация <see cref="ISectionBuilder"/>
/// </summary>
internal class SectionBuilder : OpenXmlElementBuilder, ISectionBuilder
{
	/// <summary>
	/// Настраиваемая секция
	/// </summary>
	private readonly SectionProperties _section;

	/// <summary>
	/// Последний абзац в документе
	/// </summary>
	private Paragraph LastParagraph => Body.LastOrDefaultChild<Paragraph>();

	/// <summary>
	/// Инициализирует <see cref="SectionBuilder"/> по указанному <see cref="MainDocumentPart"/>
	/// </summary>
	/// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
	/// <param name="section">Настраиваемая секция</param>
	internal SectionBuilder(MainDocumentPart mainDocumentPart, SectionProperties section)
		: base(mainDocumentPart)
	{
		_section = section;
	}

	/// <summary>
	/// Настраивает параметры страницы
	/// </summary>
	/// <param name="configurePageLayout">Метод, настраивающий параметры страницы</param>
	public ISectionBuilder ConfigurePageLayout(Action<IPageLayoutBuilder> configurePageLayout)
	{
		var pageSize = _section.FirstOrNewChild<PageSize>();
		var pageMargin = _section.FirstOrNewChild<PageMargin>();
		
		ConfigureWith<PageLayoutBuilder>(configurePageLayout, pageSize, pageMargin);
		
		return this;
	}

	/// <inheritdoc/>
	public ISectionBuilder ConfigureLastParagraph(Action<IParagraphBuilder> configureParagraph)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc/>
	public ISectionBuilder AppendParagraph(Action<IParagraphBuilder> configureParagraph)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc/>
	public ISectionBuilder AppendChunk(Action<string> configureRelationship)
	{
		var identificator = Guid.NewGuid().ToString();
		configureRelationship(identificator);
		
		var chunk = new AltChunk()
		{
			Id = identificator
		};

		Body.InsertAfter(chunk, LastParagraph);

		return this;
	}
}
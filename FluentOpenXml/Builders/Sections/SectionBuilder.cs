using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.PageLayout;
using FluentOpenXml.Builders.PageLayout.Interfaces;
using FluentOpenXml.Builders.Paragraphs;
using FluentOpenXml.Builders.Paragraphs.Interfaces;
using FluentOpenXml.Builders.Sections.Interfaces;
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
	private Paragraph LastParagraph
	{
		get
		{
			var paragraph = Body.LastOrDefaultChild<Paragraph>();
			
			ThrowIfElementNotAdded(paragraph);
			
			return paragraph;
		}
	}

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

	/// <inheritdoc/>
	public ISectionBuilder ForPageLayout(Action<IPageLayoutBuilder> configurePageLayout)
	{
		ArgumentNullException.ThrowIfNull(configurePageLayout);
		
		var pageSize = _section.FirstOrNewChild<PageSize>();
		var pageMargin = _section.FirstOrNewChild<PageMargin>();
		
		ConfigureWith<PageLayoutBuilder>(configurePageLayout, pageSize, pageMargin);
		
		return this;
	}

	/// <inheritdoc/>
	public ISectionBuilder ForLastParagraph(Action<IParagraphBuilder> configureParagraph)
	{
		ArgumentNullException.ThrowIfNull(configureParagraph);
		ConfigureWith<ParagraphBuilder>(configureParagraph, LastParagraph);
		
		return this;
	}

	/// <inheritdoc/>
	public ISectionBuilder AppendParagraph(Action<IParagraphBuilder> configureParagraph)
	{
		ArgumentNullException.ThrowIfNull(configureParagraph);

		Body.AppendChild<Paragraph>();
		return ForLastParagraph(configureParagraph);
	}

	/// <inheritdoc/>
	public ISectionBuilder AppendChunk(Action<string> configureRelationship)
	{
		ArgumentNullException.ThrowIfNull(configureRelationship);
		
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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Paragraphs.Enums;
using FluentOpenXml.Builders.Paragraphs.Interfaces;

namespace FluentOpenXml.Builders.Paragraphs;

/// <summary>
/// Стандартная реализация <see cref="IParagraphBuilder"/>
/// </summary>
internal class ParagraphBuilder : OpenXmlElementBuilder, IParagraphBuilder
{
	/// <summary>
	/// Настраиваемый абзац
	/// </summary>
	private readonly Paragraph _paragraph;

	/// <summary>
	/// Инициализирует <see cref="ParagraphBuilder"/> по указанному <see cref="MainDocumentPart"/>
	/// </summary>
	/// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
	/// <param name="paragraph">Настраиваемый абзац</param>
	public ParagraphBuilder(MainDocumentPart mainDocumentPart, Paragraph paragraph)
		: base(mainDocumentPart)
	{
		_paragraph = paragraph;
	}
	
	/// <inheritdoc />
	public IParagraphBuilder AppendText(Action<ITextBuilder> configureText)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc />
	public IParagraphBuilder SetJustification(ParagraphJustification justification)
	{
		throw new NotImplementedException();
	}

	
}
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Paragraphs.Enums;
using FluentOpenXml.Builders.Paragraphs.Interfaces;
using FluentOpenXml.Extensions;
using FluentOpenXml.Units.Universal;

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
	/// Параметры абзаца
	/// </summary>
	private ParagraphProperties Properties
	{
		get
		{
			var firstRun = _paragraph.FirstOrDefaultChild<Run>();
			return _paragraph.FirstOrPrependChild<ParagraphProperties>(firstRun);
		}
	}

	/// <summary>
	/// Последний область текста в абзаце
	/// </summary>
	private Run LastRun
	{
		get
		{
			var run = _paragraph.LastOrDefaultChild<Run>();
			
			ThrowIfElementNotAdded(run);

			return run;
		}
	}

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

	/// <summary>
	/// Вставляет макет текстовой области в абзац
	/// </summary>
	/// <returns></returns>
	private IParagraphBuilder AppendRun() => AppendRun(_ =>
	{
		// Ignore...
	});
	
	/// <inheritdoc />
	public IParagraphBuilder AppendRun(Action<IRunBuilder> configureRun)
	{
		ArgumentNullException.ThrowIfNull(configureRun);

		_paragraph.AppendChild<Run>();
		ConfigureWith<RunBuilder>(configureRun, LastRun);
		
		return this;
	}

	/// <inheritdoc />
	public IParagraphBuilder SetJustification(ParagraphJustification justification)
	{
		var justificationProperties = Properties.FirstOrNewChild<Justification>();
		justificationProperties.Val = (JustificationValues)justification;

		return this;
	}

	/// <inheritdoc />
	public IParagraphBuilder AppendPageBreak()
	{
		AppendRun();
		
		var pageBreak = LastRun.FirstOrNewChild<Break>();
		@pageBreak.Type = BreakValues.Page;

		return this;
	}

	/// <inheritdoc />
	public IParagraphBuilder SpacingBefore<TUnits>(double spacingBefore) 
		where TUnits : UniversalUnits
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc />
	public IParagraphBuilder SpacingAfter<TUnits>(double spacingAfter) 
		where TUnits : UniversalUnits
	{
		throw new NotImplementedException();
	}
}
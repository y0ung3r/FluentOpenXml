using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Paragraphs.Interfaces;
using Color = System.Drawing.Color;

namespace FluentOpenXml.Builders.Paragraphs;

/// <summary>
/// Стандартная реализация <see cref="IRunBuilder"/>
/// </summary>
internal class RunBuilder : OpenXmlElementBuilder, IRunBuilder
{
	/// <summary>
	/// Настраиваемая область текста
	/// </summary>
	private readonly Run _run;

	/// <summary>
	/// Инициализирует <see cref="RunBuilder"/> по указанному <see cref="MainDocumentPart"/>
	/// </summary>
	/// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
	/// <param name="run">Настраиваемая область текста</param>
	public RunBuilder(MainDocumentPart mainDocumentPart, Run run) 
		: base(mainDocumentPart)
	{
		_run = run;
	}
	
	/// <inheritdoc />
	public IRunBuilder SetText(string text)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc />
	public IRunBuilder SetFontFamily(string fontFamily)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc />
	public IRunBuilder SetFontSize(double fontSize)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc />
	public IRunBuilder SetColor(Color color)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc />
	public IRunBuilder IsBold(bool isBold)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc />
	public IRunBuilder IsItalic(bool isItalic)
	{
		throw new NotImplementedException();
	}
}
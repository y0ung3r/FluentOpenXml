using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Paragraphs.Interfaces;
using FluentOpenXml.Builders.Sections.Enums;
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
	private SectionProperties _section;

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
	/// Устанавливает указанную ориентацию для страницы
	/// </summary>
	/// <param name="orientation">Ориентация</param>
	public ISectionBuilder SetPageOrientation(PageOrientation orientation)
	{
		throw new NotImplementedException();
	}

	/// <summary>
	/// Устанавливает размеры страницы в соответствии с заданными параметры
	/// </summary>
	/// <param name="applyPageSize">Настраивает размеры страницы</param>
	public ISectionBuilder SetPageSize(Action<IPageSizeBuilder> applyPageSize)
	{
		throw new NotImplementedException();
	}

	/// <summary>
	/// Устанавливает внешние отступы у страницы в соответствии с заданными параметрами
	/// </summary>
	/// <param name="applyPageMargin">Настраивает внешние границы</param>
	public ISectionBuilder SetPageMargin(Action<IPageMarginBuilder> applyPageMargin)
	{
		throw new NotImplementedException();
	}

	/// <summary>
	/// Возвращает последний абзац
	/// </summary>
	/// <param name="configureParagraph">Настраивает абзац</param>
	public ISectionBuilder ConfigureLastParagraph(Action<IParagraphBuilder> configureParagraph)
	{
		throw new NotImplementedException();
	}

	/// <summary>
	/// Вставляет абзац в конец секции
	/// </summary>
	/// <param name="configureParagraph">Настраивает абзац</param>
	/// <returns></returns>
	public ISectionBuilder AppendParagraph(Action<IParagraphBuilder> configureParagraph)
	{
		throw new NotImplementedException();
	}

	/// <summary>
	/// Вставляет якорь с указанным уникальным идентификатором для вставки внешнего документа
	/// </summary>
	/// <param name="identificator">Уникальный идентификатор</param>
	public ISectionBuilder AppendChunk(string identificator)
	{
		var chunk = new AltChunk()
		{
			Id = identificator
		};

		Body.InsertAfter(chunk, LastParagraph);

		return this;
	}
}
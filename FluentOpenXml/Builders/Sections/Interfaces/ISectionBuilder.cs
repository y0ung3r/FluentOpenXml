using FluentOpenXml.Builders.Paragraphs.Interfaces;
using FluentOpenXml.Builders.Sections.Enums;
using FluentOpenXml.Builders.Sections.PageLayout.Interfaces;

namespace FluentOpenXml.Builders.Sections.Interfaces;

/// <summary>
/// Представляет построитель секции в документе
/// </summary>
public interface ISectionBuilder
{
	/// <summary>
	/// Настраивает параметры страницы
	/// </summary>
	/// <param name="configurePageLayout">Метод, настраивающий параметры страницы</param>
	ISectionBuilder ConfigurePageLayout(Action<IPageLayoutBuilder> configurePageLayout);

	/// <summary>
	/// Настраивает последний абзац
	/// </summary>
	/// <param name="configureParagraph">Метод, настраивающий абзац</param>
	ISectionBuilder ConfigureLastParagraph(Action<IParagraphBuilder> configureParagraph);

	/// <summary>
	/// Вставляет абзац в конец секции
	/// </summary>
	/// <param name="configureParagraph">Настраивает абзац</param>
	ISectionBuilder AppendParagraph(Action<IParagraphBuilder> configureParagraph);

	/// <summary>
	/// Вставляет якорь для вставки внешнего документа
	/// </summary>
	/// <param name="configureRelationship">Настраивает вставку внешнего документа, используя уникальный идентификатор</param>
	ISectionBuilder AppendChunk(Action<string> configureRelationship);
}
using FluentOpenXml.Builders.Paragraphs.Interfaces;
using FluentOpenXml.Builders.Sections.Enums;

namespace FluentOpenXml.Builders.Sections.Interfaces;

/// <summary>
/// Представляет построитель секции в документе
/// </summary>
public interface ISectionBuilder
{
	/// <summary>
	/// Устанавливает указанную ориентацию для страницы
	/// </summary>
	/// <param name="orientation">Ориентация</param>
	ISectionBuilder SetPageOrientation(PageOrientation orientation);

	/// <summary>
	/// Устанавливает размеры страницы в соответствии с заданными параметры
	/// </summary>
	/// <param name="setPageSize">Настраивает размеры страницы</param>
	ISectionBuilder SetPageSize(Action<IPageSizeBuilder> setPageSize);

	/// <summary>
	/// Устанавливает внешние отступы у страницы в соответствии с заданными параметрами
	/// </summary>
	/// <param name="setPageMargin">Настраивает внешние границы</param>
	ISectionBuilder SetPageMargin(Action<IPageMarginBuilder> setPageMargin);

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
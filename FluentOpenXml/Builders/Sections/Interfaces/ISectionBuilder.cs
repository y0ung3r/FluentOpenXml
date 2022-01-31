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
    /// <param name="applyPageSize">Настраивает размеры страницы</param>
    ISectionBuilder SetPageSize(Action<IPageSizeBuilder> applyPageSize);

    /// <summary>
    /// Устанавливает внешние отступы у страницы в соответствии с заданными параметрами
    /// </summary>
    /// <param name="applyPageMargin">Настраивает внешние границы</param>
    ISectionBuilder SetPageMargin(Action<IPageMarginBuilder> applyPageMargin);

    /// <summary>
    /// Вставляет абзац в конец секции
    /// </summary>
    /// <param name="configureParagraph">Настраивает абзац</param>
    /// <returns></returns>
    ISectionBuilder AppendParagraph(Action<IParagraphBuilder> configureParagraph);
}
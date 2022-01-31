using FluentOpenXml.Builders.Section.Enums;

namespace FluentOpenXml.Builders.Section.Interfaces;

/// <summary>
/// Представляет построитель секции в документе
/// </summary>
public interface ISectionBuilder
{
    /// <summary>
    /// Устанавливает указанную ориентацию для страницы
    /// </summary>
    /// <param name="orientation">Ориентация</param>
    ISectionBuilder SetPageOrientation(PageOrientationType orientation);

    /// <summary>
    /// Устанавливает размеры страницы в соответствии с заданными параметры
    /// </summary>
    /// <param name="applyPageSize">Построитель параметров</param>
    ISectionBuilder SetPageSize(Action<IPageSizeBuilder> applyPageSize);
}
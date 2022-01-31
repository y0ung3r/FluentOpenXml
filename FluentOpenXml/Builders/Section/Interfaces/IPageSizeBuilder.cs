namespace FluentOpenXml.Builders.Section.Interfaces;

/// <summary>
/// Представляет построить размеров страницы
/// </summary>
public interface IPageSizeBuilder
{
    /// <summary>
    /// Устанавливает ширину страницы
    /// </summary>
    /// <param name="value">Ширина страницы</param>
    IPageSizeBuilder SetWidth(double value);

    /// <summary>
    /// Устанавливает высоту страницы
    /// </summary>
    /// <param name="value">Высота страницы</param>
    IPageSizeBuilder SetHeight(double value);
}
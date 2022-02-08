namespace FluentOpenXml.Builders.Sections.Interfaces;

/// <summary>
/// Представляет построитель размеров страницы
/// Внутри Open XML задается в Dxa
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
using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.Sections.Interfaces;

/// <summary>
/// Представляет построитель размеров страницы
/// </summary>
public interface IPageSizeBuilder
{
    /// <summary>
    /// Устанавливает ширину страницы
    /// </summary>
    /// <param name="value">Ширина страницы</param>
    IPageSizeBuilder SetWidth(UniversalUnits value);

    /// <summary>
    /// Устанавливает высоту страницы
    /// </summary>
    /// <param name="value">Высота страницы</param>
    IPageSizeBuilder SetHeight(UniversalUnits value);
}
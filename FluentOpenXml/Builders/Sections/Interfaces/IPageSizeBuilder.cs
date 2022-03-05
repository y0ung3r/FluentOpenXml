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
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageSizeBuilder SetWidth<TUnits>(TUnits value)
        where TUnits : UniversalUnits;

    /// <summary>
    /// Устанавливает высоту страницы
    /// </summary>
    /// <param name="value">Высота страницы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageSizeBuilder SetHeight<TUnits>(TUnits value)
        where TUnits : UniversalUnits;
}
using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.Sections.Interfaces;

/// <summary>
/// Представляет построитель размеров страницы
/// </summary>
public interface IPageSizeBuilder
{
    /// <summary>
    /// Возвращает ширину страницы в указанных единицах измерения
    /// </summary>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    TUnits GetWidth<TUnits>()
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает ширину страницы
    /// </summary>
    /// <param name="value">Ширина страницы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageSizeBuilder SetWidth<TUnits>(double value)
        where TUnits : UniversalUnits;

    /// <summary>
    /// Возвращает высоту страницы в указанных единицах измерения
    /// </summary>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    TUnits GetHeight<TUnits>()
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает высоту страницы
    /// </summary>
    /// <param name="value">Высота страницы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageSizeBuilder SetHeight<TUnits>(double value)
        where TUnits : UniversalUnits;
}
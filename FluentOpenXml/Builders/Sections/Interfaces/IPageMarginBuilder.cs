using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.Sections.Interfaces;

/// <summary>
/// Представляет построитель внешних отступов у страницы
/// </summary>
public interface IPageMarginBuilder
{
    /// <summary>
    /// Устанавливает левую внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от левой границы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetLeft<TUnits>(TUnits value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает верхнюю внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от верхней границы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetTop<TUnits>(TUnits value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает правую внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от правой границы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetRight<TUnits>(TUnits value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает нижнюю внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от нижней границы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetBottom<TUnits>(TUnits value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает отступ от верхнего колонтитула
    /// </summary>
    /// <param name="value">Внешний отступ от верхнего колонтитула</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetHeader<TUnits>(TUnits value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает отступ от нижнего колонтитула
    /// </summary>
    /// <param name="value">Внешний отступ от нижнего колонтитула</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetFooter<TUnits>(TUnits value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает интервалы переплета страниц
    /// </summary>
    /// <param name="value">Интервал переплета страниц</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetGutter<TUnits>(TUnits value)
        where TUnits : UniversalUnits;
}
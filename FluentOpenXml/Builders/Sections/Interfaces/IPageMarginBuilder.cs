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
    IPageMarginBuilder SetLeft(UniversalUnits value);
    
    /// <summary>
    /// Устанавливает верхнюю внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от верхней границы</param>
    IPageMarginBuilder SetTop(UniversalUnits value);
    
    /// <summary>
    /// Устанавливает правую внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от правой границы</param>
    IPageMarginBuilder SetRight(UniversalUnits value);
    
    /// <summary>
    /// Устанавливает нижнюю внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от нижней границы</param>
    IPageMarginBuilder SetBottom(UniversalUnits value);
    
    /// <summary>
    /// Устанавливает отступ от верхнего колонтитула
    /// </summary>
    /// <param name="value">Внешний отступ от верхнего колонтитула</param>
    IPageMarginBuilder SetHeader(UniversalUnits value);
    
    /// <summary>
    /// Устанавливает отступ от нижнего колонтитула
    /// </summary>
    /// <param name="value">Внешний отступ от нижнего колонтитула</param>
    IPageMarginBuilder SetFooter(UniversalUnits value);
    
    /// <summary>
    /// Устанавливает интервалы переплета страниц
    /// </summary>
    /// <param name="value">Интервал переплета страниц</param>
    IPageMarginBuilder SetGutter(UniversalUnits value);
}
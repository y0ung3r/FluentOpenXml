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
    IPageMarginBuilder SetLeft(float value);
    
    /// <summary>
    /// Устанавливает верхнюю внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от верхней границы</param>
    IPageMarginBuilder SetTop(float value);
    
    /// <summary>
    /// Устанавливает правую внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от правой границы</param>
    IPageMarginBuilder SetRight(float value);
    
    /// <summary>
    /// Устанавливает нижнюю внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от нижней границы</param>
    IPageMarginBuilder SetBottom(float value);
    
    /// <summary>
    /// Устанавливает отступ от верхнего колонтитула
    /// </summary>
    /// <param name="value">Внешний отступ от верхнего колонтитула</param>
    IPageMarginBuilder SetHeader(float value);
    
    /// <summary>
    /// Устанавливает отступ от нижнего колонтитула
    /// </summary>
    /// <param name="value">Внешний отступ от нижнего колонтитула</param>
    IPageMarginBuilder SetFooter(float value);
    
    /// <summary>
    /// Устанавливает интервалы переплета страниц
    /// </summary>
    /// <param name="value">Интервал переплета страниц</param>
    IPageMarginBuilder SetGutter(float value);
}
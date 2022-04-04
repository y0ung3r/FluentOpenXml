﻿using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.PageLayout.Interfaces;

/// <summary>
/// Представляет построитель внешних отступов у страницы
/// </summary>
public interface IPageMarginBuilder : ITransposable
{
    /// <summary>
    /// Устанавливает левую внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от левой границы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetLeft<TUnits>(double value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает верхнюю внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от верхней границы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetTop<TUnits>(double value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает правую внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от правой границы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetRight<TUnits>(double value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает нижнюю внешнюю границу
    /// </summary>
    /// <param name="value">Внешний отступ от нижней границы</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetBottom<TUnits>(double value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает отступ от верхнего колонтитула
    /// </summary>
    /// <param name="value">Внешний отступ от верхнего колонтитула</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetHeader<TUnits>(double value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает отступ от нижнего колонтитула
    /// </summary>
    /// <param name="value">Внешний отступ от нижнего колонтитула</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetFooter<TUnits>(double value)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает интервалы переплета страниц
    /// </summary>
    /// <param name="value">Интервал переплета страниц</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IPageMarginBuilder SetGutter<TUnits>(double value)
        where TUnits : UniversalUnits;
}
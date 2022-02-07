// ReSharper disable InconsistentNaming

namespace FluentOpenXml.Units.Enums;

/// <summary>
/// Представляет типы размеров
/// </summary>
public enum UnitType
{
    /// <summary>
    /// Английские единицы измерения
    /// https://en.wikipedia.org/wiki/English_units
    /// https://en.wikipedia.org/wiki/Office_Open_XML_file_formats
    /// </summary>
    Emu,
    
    /// <summary>
    /// Пиксели
    /// </summary>
    Pixels,
    
    /// <summary>
    /// Пятидесятые доли процента (пункты)
    /// Используется для относительных измерений: ширина таблицы, ячейки, поля страницы и т. д.
    /// </summary>
    Pcts,
    
    /// <summary>
    /// Проценты
    /// </summary>
    Percentages,
    
    /// <summary>
    /// Двадцатые доли точки
    /// Используется в отступах, интервалах и т. д.
    /// </summary>
    Dxa,
    
    /// <summary>
    /// Точки
    /// </summary>
    Points,
    
    /// <summary>
    /// Половины от точки
    /// Используется в размерах шрифта
    /// </summary>
    HalfPoints,

    /// <summary>
    /// Дюймы
    /// </summary>
    Inches,
    
    /// <summary>
    /// Миллиметры
    /// </summary>
    Millimeters,
    
    /// <summary>
    /// Сантиметры
    /// </summary>
    Centimeters
}
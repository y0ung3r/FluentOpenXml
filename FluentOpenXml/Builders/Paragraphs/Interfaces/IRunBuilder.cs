using System.Drawing;
using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.Paragraphs.Interfaces;

/// <summary>
/// Представляет построитель текстовой области
/// </summary>
public interface IRunBuilder
{
    /// <summary>
    /// Устанавливает текст
    /// </summary>
    /// <param name="text">Текст</param>
    IRunBuilder SetText(string text);
    
    /// <summary>
    /// Устанавливает шрифт (например, "Arial")
    /// </summary>
    /// <param name="fontFamily">Наименование шрифта</param>
    IRunBuilder SetFontFamily(string fontFamily);

    /// <summary>
    /// Устанавливает размер шрифта
    /// </summary>
    /// <param name="fontSize">Размер шрифта в <see cref="Points"/></param>
    IRunBuilder SetFontSize(double fontSize);

    /// <summary>
    /// Устанавливает цвет
    /// </summary>
    /// <param name="color">Цвет</param>
    IRunBuilder SetColor(Color color);

    /// <summary>
    /// Выделяет текст полужирным
    /// </summary>
    /// <param name="isBold">True - текст полужирный, False - обычный текст</param>
    IRunBuilder IsBold(bool isBold);

    /// <summary>
    /// Наклоняет текст
    /// </summary>
    /// <param name="isItalic">True - наклоненный текст, False - обычный текст</param>
    IRunBuilder IsItalic(bool isItalic);
}
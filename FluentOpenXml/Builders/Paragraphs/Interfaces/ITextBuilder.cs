using System.Drawing;
using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.Paragraphs.Interfaces;

/// <summary>
/// Представляет построитель текста
/// </summary>
public interface ITextBuilder
{
    /// <summary>
    /// Устанавливает шрифт (например, "Arial")
    /// </summary>
    /// <param name="fontFamily">Наименование шрифта</param>
    ITextBuilder SetFontFamily(string fontFamily);

    /// <summary>
    /// Устанавливает размер шрифта
    /// </summary>
    /// <param name="fontSize">Размер шрифта в <see cref="Points"/></param>
    ITextBuilder SetFontSize(double fontSize);

    /// <summary>
    /// Устанавливает цвет
    /// </summary>
    /// <param name="color">Цвет</param>
    ITextBuilder SetColor(Color color);

    /// <summary>
    /// Выделяет текст полужирным
    /// </summary>
    /// <param name="isBold">True - текст полужирный, False - обычный текст</param>
    ITextBuilder IsBold(bool isBold);

    /// <summary>
    /// Наклоняет текст
    /// </summary>
    /// <param name="isItalic">True - наклоненный текст, False - обычный текст</param>
    ITextBuilder IsItalic(bool isItalic);
}
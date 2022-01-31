using System.Drawing;

namespace FluentOpenXml.Builders.Paragraphs.Interfaces;

/// <summary>
/// Представляет построитель текста
/// </summary>
public interface ITextBuilder
{
    /// <summary>
    /// Устанавливает шрифт
    /// </summary>
    /// <param name=""></param>
    ITextBuilder SetFont();

    /// <summary>
    /// Устанавливает размер шрифта
    /// </summary>
    ITextBuilder SetFontSize(float fontSize);

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
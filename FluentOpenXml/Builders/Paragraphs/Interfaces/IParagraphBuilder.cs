using FluentOpenXml.Builders.Paragraphs.Enums;

namespace FluentOpenXml.Builders.Paragraphs.Interfaces;

/// <summary>
/// Представляет построитель абзацев
/// </summary>
public interface IParagraphBuilder
{
    /// <summary>
    /// Вставляет текст в абзац
    /// </summary>
    /// <param name="configureText">Настраивает текст</param>
    IParagraphBuilder AppendText(Action<ITextBuilder> configureText);

    /// <summary>
    /// Устанавливает выравнивание для абзаца
    /// </summary>
    /// <param name="justification">Выравнивание</param>
    IParagraphBuilder SetJustification(ParagraphJustification justification);
}
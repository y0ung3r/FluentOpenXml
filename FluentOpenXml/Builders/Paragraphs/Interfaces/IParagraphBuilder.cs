using FluentOpenXml.Builders.Paragraphs.Enums;
using FluentOpenXml.Units.Universal;

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

    /// <summary>
    /// Вставляет разрыв страницы
    /// </summary>
    IParagraphBuilder AppendPageBreak();
    
    /// <summary>
    /// Устанавливает отступ перед абзацем
    /// </summary>
    /// <param name="spacingBefore">Значение отступа</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IParagraphBuilder SpacingBefore<TUnits>(double spacingBefore)
        where TUnits : UniversalUnits;
    
    /// <summary>
    /// Устанавливает отступ после абзаца
    /// </summary>
    /// <param name="spacingAfter">Значение отступа</param>
    /// <typeparam name="TUnits">Универсальные единицы измерения</typeparam>
    IParagraphBuilder SpacingAfter<TUnits>(double spacingAfter)
        where TUnits : UniversalUnits;
}
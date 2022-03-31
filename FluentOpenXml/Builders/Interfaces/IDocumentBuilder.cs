using FluentOpenXml.Builders.Sections.Interfaces;

namespace FluentOpenXml.Builders.Interfaces;

/// <summary>
/// Представляет построитель документа
/// </summary>
public interface IDocumentBuilder
{
    /// <summary>
    /// Настраивает последнюю секцию в документе
    /// </summary>
    /// <param name="configureSection">Метод, настраивающий секцию</param>
    IDocumentBuilder ForLastSection(Action<ISectionBuilder> configureSection);

    /// <summary>
    /// Вставляет внешний документ в конец текущего
    /// </summary>
    /// <param name="Stream">Последовательность байтов вставляемого документа</param>
    IDocumentBuilder AppendDocument(Stream stream);

    /// <summary>
    /// Вставляет новую секцию в конец документа с помощью разрыва страницы
    /// </summary>
    /// <param name="configureSection">Настраивает новую секцию</param>
    IDocumentBuilder AppendSectionBreak(Action<ISectionBuilder> configureSection);
    
    /// <summary>
    /// Очищает документ
    /// </summary>
    IDocumentBuilder Clear();
}
using FluentOpenXml.Interfaces;

namespace FluentOpenXml;

/// <summary>
/// Представляет стандартную реализацию <see cref="IDocumentLoader"/>
/// </summary>
internal class DocumentLoader : IDocumentLoader
{
    /// <summary>
    /// Поставщик сервисов
    /// </summary>
    public IServiceProvider ServiceProvider { get; }
    
    /// <summary>
    /// Инициализирует <see cref="DocumentLoader"/>
    /// </summary>
    /// <param name="serviceProvider">Поставщик сервисов</param>
    public DocumentLoader(IServiceProvider serviceProvider)
    {
        ServiceProvider = serviceProvider;
    }
}
using Microsoft.Extensions.DependencyInjection;

namespace FluentOpenXml.Interfaces;

/// <summary>
/// Определяет механизм создания/открытия Open XML документа
/// </summary>
public interface IDocumentLoader
{
    /// <summary>
    /// Поставщик сервисов
    /// </summary>
    IServiceProvider ServiceProvider { get; }

    /// <summary>
    /// Создает новый документ
    /// </summary>
    IOpenXmlDocument Create() => ActivatorUtilities.CreateInstance<OpenXmlDocument>
    (
        ServiceProvider
    );
    
    /// <summary>
    /// Создает новый документ и применяет к нему указанные настройки
    /// </summary>
    /// <param name="settings">Настройки для документа</param>
    IOpenXmlDocument Create(DocumentSettings settings) => ActivatorUtilities.CreateInstance<OpenXmlDocument>
    (
        ServiceProvider, 
        settings
    );

    /// <summary>
    /// Открывает документ с помощью потока данных
    /// </summary>
    /// <param name="stream">Поток данных</param>
    IOpenXmlDocument Open(Stream stream) => ActivatorUtilities.CreateInstance<OpenXmlDocument>
    (
        ServiceProvider, 
        stream
    );

    /// <summary>
    /// Открывает документ с помощью потока данных и применяет к нему указанные настройки
    /// </summary>
    /// <param name="stream">Поток данных</param>
    /// <param name="settings">Настройки для документа</param>
    /// <returns></returns>
    IOpenXmlDocument Open(Stream stream, DocumentSettings settings) => ActivatorUtilities.CreateInstance<OpenXmlDocument>
    (
        ServiceProvider, 
        stream, 
        settings
    );

    /// <summary>
    /// Открывает документ с помощью указанного пути
    /// </summary>
    /// <param name="filepath">Путь к документу</param>
    IOpenXmlDocument Open(string filepath) => ActivatorUtilities.CreateInstance<OpenXmlDocument>
    (
        ServiceProvider, 
        filepath
    );
    
    /// <summary>
    /// Открывает документ с помощью указанного пути и применяет к нему указанные настройки
    /// </summary>
    /// <param name="filepath">Путь к документу</param>
    /// <param name="settings">Настройки для документа</param>
    IOpenXmlDocument Open(string filepath, DocumentSettings settings) => ActivatorUtilities.CreateInstance<OpenXmlDocument>
    (
        ServiceProvider, 
        filepath,
        settings
    );
}
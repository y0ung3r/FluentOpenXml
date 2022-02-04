using FluentOpenXml.Builders;
using FluentOpenXml.Builders.Interfaces;
using FluentOpenXml.Builders.Sections;
using FluentOpenXml.Builders.Sections.Interfaces;
using FluentOpenXml.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace FluentOpenXml.Extensions;

/// <summary>
/// Методы-расширения для <see cref="IServiceProvider"/>
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Добавляет FluentOpenXml в контейнер зависимостей
    /// </summary>
    /// <param name="services">Контейнер зависимостей</param>
    public static IServiceCollection AddFluentOpenXml(this IServiceCollection services)
    {
        services.TryAddTransient<IDocumentLoader>();
        services.TryAddTransient<IDocumentBuilder, DocumentBuilder>();
        services.TryAddTransient<ISectionBuilder, SectionBuilder>();
        
        return services;
    }
}
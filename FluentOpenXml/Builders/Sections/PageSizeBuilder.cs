using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Sections.Interfaces;

namespace FluentOpenXml.Builders.Sections;

/// <summary>
/// Стандартная реализация <see cref="IPageSizeBuilder"/>
/// </summary>
internal class PageSizeBuilder : OpenXmlElementBuilder, IPageSizeBuilder
{
    /// <summary>
    /// Размеры страницы
    /// </summary>
    private PageSize _pageSize;

    /// <summary>
    /// Инициализирует <see cref="PageSizeBuilder"/> по указанному <see cref="MainDocumentPart"/> и <see cref="PageSize"/>
    /// </summary>
    /// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
    /// <param name="pageSize">Размеры страницы</param>
    public PageSizeBuilder(MainDocumentPart mainDocumentPart, PageSize pageSize) 
        : base(mainDocumentPart)
    {
        _pageSize = pageSize;
    }
    
    /// <inheritdoc />
    public IPageSizeBuilder SetWidth(double value)
    {
        if (value < 0.0)
        {
            throw new ArgumentException($"Значение \"{nameof(value)}\" должно быть больше нуля");
        }
        
        // _pageSize.Width = value;

        return this;
    }

    /// <inheritdoc />
    public IPageSizeBuilder SetHeight(double value)
    {
        throw new NotImplementedException();
    }
}
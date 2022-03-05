using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Sections.Interfaces;
using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.Sections;

/// <summary>
/// Стандартная реализация <see cref="IPageSizeBuilder"/>
/// </summary>
internal class PageSizeBuilder : OpenXmlElementBuilder, IPageSizeBuilder
{
    /// <summary>
    /// Размеры страницы
    /// </summary>
    private readonly PageSize _pageSize;

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
    public IPageSizeBuilder SetWidth<TUnits>(TUnits value)
        where TUnits : UniversalUnits
    {
        SetPropertyValue
        (
            _pageSize, size => size.Width, 
            value, units => units.ToEmu().ToTwips()
        );

        return this;
    }

    /// <inheritdoc />
    public IPageSizeBuilder SetHeight<TUnits>(TUnits value)
        where TUnits : UniversalUnits
    {
        SetPropertyValue
        (
            _pageSize, size => size.Height, 
            value, units => units.ToEmu().ToTwips()
        );

        return this;
    }
}
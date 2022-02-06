using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Interfaces;
using FluentOpenXml.Builders.Sections;
using FluentOpenXml.Builders.Sections.Interfaces;
using FluentOpenXml.Extensions;

namespace FluentOpenXml.Builders;

/// <summary>
/// Стандартная реализация для <see cref="IDocumentBuilder"/>
/// </summary>
internal class DocumentBuilder : OpenXmlElementBuilder, IDocumentBuilder
{
    /// <summary>
    /// Последняя секция в документе
    /// </summary>
    private SectionProperties LastSection => Body.FirstOrDefaultChild<SectionProperties>();

    /// <summary>
    /// Инициализирует <see cref="DocumentBuilder"/> по указанному <see cref="MainDocumentPart"/>
    /// </summary>
    /// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
    internal DocumentBuilder(MainDocumentPart mainDocumentPart)
        : base(mainDocumentPart)
    { }

    /// <summary>
    /// Конфигурирует указанную секцию в документе
    /// </summary>
    /// <param name="section">Секция</param>
    /// <param name="configureSection">Настраивает секцию</param>
    private IDocumentBuilder ConfigureSection(SectionProperties section, Action<ISectionBuilder> configureSection)
    {
        ArgumentNullException.ThrowIfNull(section);
        ArgumentNullException.ThrowIfNull(configureSection);
        
        var sectionBuilder = new SectionBuilder(MainDocumentPart, section);
        configureSection(sectionBuilder);

        return this;
    }

    /// <inheritdoc/>
    public IDocumentBuilder ConfigureLastSection(Action<ISectionBuilder> configureSection) => ConfigureSection
    (
        LastSection,
        configureSection
    );

    /// <inheritdoc/>
    public IDocumentBuilder AppendAnotherDocument(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        
        var relationshipId = Guid.NewGuid().ToString();
        var partType = AlternativeFormatImportPartType.WordprocessingML;

        var part = MainDocumentPart.AddAlternativeFormatImportPart(partType, relationshipId);
        part.FeedData(stream);

        return ConfigureLastSection
        (
            section => section.AppendChunk(relationshipId)
        );
    }

    /// <inheritdoc/>
    public IDocumentBuilder AppendSectionBreak(Action<ISectionBuilder> configureSection)
    {
        ArgumentNullException.ThrowIfNull(configureSection);
        
        var sectionType = LastSection.FirstOrNewChild<SectionType>();
        sectionType.Val = SectionMarkValues.NextPage;

        var section = Body.AppendChild<SectionProperties>();
        return ConfigureSection(section, configureSection);
    }

    /// <inheritdoc/>
    public IDocumentBuilder Clear()
    {
        Body.RemoveAllChildren();
        
        return this;
    }
}
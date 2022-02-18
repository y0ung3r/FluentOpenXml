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

    /// <inheritdoc/>
    public IDocumentBuilder ConfigureLastSection(Action<ISectionBuilder> configureSection)
    {
        Configure<SectionBuilder, SectionProperties>(LastSection, configureSection);

        return this;
    }

    /// <inheritdoc/>
    public IDocumentBuilder AppendAnotherDocument(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

        return ConfigureLastSection
        (
            section => section.AppendChunk
            (
                relationshipId =>
                {
                    var partType = AlternativeFormatImportPartType.WordprocessingML;
                    var part = MainDocumentPart.AddAlternativeFormatImportPart(partType, relationshipId);
                    part.FeedData(stream);
                }
            )
        );
    }

    /// <inheritdoc/>
    public IDocumentBuilder AppendSectionBreak(Action<ISectionBuilder> configureSection)
    {
        ArgumentNullException.ThrowIfNull(configureSection);
        
        var sectionType = LastSection.FirstOrNewChild<SectionType>();
        sectionType.Val = SectionMarkValues.NextPage;

        var section = Body.AppendChild<SectionProperties>();
        Configure<SectionBuilder, SectionProperties>(section, configureSection);

        return this;
    }

    /// <inheritdoc/>
    public IDocumentBuilder Clear()
    {
        Body.RemoveAllChildren();
        
        return this;
    }
}
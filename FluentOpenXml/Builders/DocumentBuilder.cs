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
    private SectionProperties LastSection
    {
        get
        {
            var section = Body.FirstOrDefaultChild<SectionProperties>();

            ThrowIfElementNotAdded(section);
            
            return section;
        }
    }

    /// <summary>
    /// Инициализирует <see cref="DocumentBuilder"/> по указанному <see cref="MainDocumentPart"/>
    /// </summary>
    /// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
    internal DocumentBuilder(MainDocumentPart mainDocumentPart)
        : base(mainDocumentPart)
    { }

    /// <inheritdoc/>
    public IDocumentBuilder ForLastSection(Action<ISectionBuilder> configureSection)
    {
        ArgumentNullException.ThrowIfNull(configureSection);
        
        ConfigureWith<SectionBuilder>(configureSection, LastSection);

        return this;
    }

    /// <inheritdoc/>
    public IDocumentBuilder AppendDocument(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

        return ForLastSection
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

        Body.AppendChild<SectionProperties>();
        return ForLastSection(configureSection);
    }

    /// <inheritdoc/>
    public IDocumentBuilder Clear()
    {
        Body.RemoveAllChildren();
        
        return this;
    }
}
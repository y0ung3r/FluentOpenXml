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
    /// Возвращает последнюю секцию в документе
    /// </summary>
    /// <param name="configureSection">Настраивает секцию</param>
    public IDocumentBuilder ConfigureLastSection(Action<ISectionBuilder> configureSection)
    {
        var sectionBuilder = new SectionBuilder(MainDocumentPart, LastSection);
        configureSection(sectionBuilder);

        return this;
    }

    /// <summary>
    /// Вставляет внешний документ в конец текущего
    /// </summary>
    /// <param name="Stream">Последовательность байтов вставляемого документа</param>
    public IDocumentBuilder AppendAnotherDocument(Stream stream)
    {
        var relationshipId = Guid.NewGuid().ToString();
        var partType = AlternativeFormatImportPartType.WordprocessingML;

        var part = MainDocumentPart.AddAlternativeFormatImportPart(partType, relationshipId);
        part.FeedData(stream);

        return ConfigureLastSection
        (
            section => section.AppendChunk(relationshipId)
        );
    }

    /// <summary>
    /// Вставляет новую секцию в конец документа с помощью разрыва страницы
    /// </summary>
    /// <param name="configureSection">Настраивает новую секцию</param>
    public IDocumentBuilder AppendSectionBreak(Action<ISectionBuilder> configureSection)
    {
        var sectionType = LastSection.FirstOrNewChild<SectionType>();
        sectionType.Val = SectionMarkValues.NextPage;

        var section = Body.AppendChild<SectionProperties>();

        var sectionBuilder = new SectionBuilder(MainDocumentPart, section);
        configureSection(sectionBuilder);

        return this;
    }

    /// <summary>
    /// Очищает документ
    /// </summary>
    public IDocumentBuilder Clear()
    {
        Body.RemoveAllChildren();
        
        return this;
    }
}
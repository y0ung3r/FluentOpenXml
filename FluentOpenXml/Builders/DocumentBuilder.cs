using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Interfaces;
using FluentOpenXml.Builders.Sections.Interfaces;
using FluentOpenXml.Extensions;

namespace FluentOpenXml.Builders;

/// <summary>
/// Стандартная реализация для <see cref="IDocumentBuilder"/>
/// </summary>
internal class DocumentBuilder : IDocumentBuilder
{
    /// <summary>
    /// Основной пакет документа OpenXML
    /// </summary>
    private readonly MainDocumentPart _mainDocumentPart;

    /// <summary>
    /// Тело документа
    /// </summary>
    private Body Body => _mainDocumentPart.Document.Body;
    
    /// <summary>
    /// Последняя секция в документе
    /// </summary>
    internal SectionProperties LastSection => Body.FirstOrDefaultChild<SectionProperties>();
    
    /// <summary>
    /// Инициализирует <see cref="DocumentBuilder"/> по указанному <see cref="Body"/>
    /// </summary>
    /// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
    internal DocumentBuilder(MainDocumentPart mainDocumentPart)
    {
        _mainDocumentPart = mainDocumentPart;
    }
    
    /// <summary>
    /// Возвращает последнюю секцию в документе
    /// </summary>
    /// <param name="configureSection">Настраивает секцию</param>
    public IDocumentBuilder GetLastSection(Action<ISectionBuilder> configureSection)
    {
        // configureSection(new SectionBuilder(section));
        
        return this;
    }

    /// <summary>
    /// Вставляет другой документ в конец текущего
    /// </summary>
    /// <param name="Stream">Последовательность байтов вставляемого документа</param>
    public IDocumentBuilder AppendAnotherDocument(Stream stream)
    {
        // Генерируем идентификатор точки в документе, в которую будет вставлен документ
        var relationshipId = Guid.NewGuid().ToString();
        
        _mainDocumentPart.AddAlternativeFormatImportPart
        (
            AlternativeFormatImportPartType.WordprocessingML, 
            relationshipId
        )
        .FeedData(stream);
        
        // Создаем абзац, после которого будет произведена вставка
        var relationshipParagraph = Body.LastOrDefaultChild<Paragraph>();
       
        // Осуществляем вставку с помощью чанка, используя созданную ранее точку в документе и новый абзац
        Body.InsertAfter
        (
            new AltChunk()
            {
                Id = relationshipId
            }, 
            relationshipParagraph
        );
        
        return this;
    }

    /// <summary>
    /// Вставляет новую секцию в конец документа с помощью разрыва и возвращает ее
    /// </summary>
    /// <param name="configureSection">Настраивает новую секцию</param>
    public IDocumentBuilder AppendSectionBreak(Action<ISectionBuilder> configureSection)
    {
        var sectionType = LastSection.FirstOrNewChild<SectionType>();
        sectionType.Val = SectionMarkValues.NextPage;

        var section = Body.AppendChild<SectionProperties>();
        // configureSection(new SectionBuilder(section));
        
        return this;
    }
}
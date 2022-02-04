using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FluentOpenXml.Builders;

/// <summary>
/// Представляет абстрактный построитель элементов Open XML
/// </summary>
internal abstract class OpenXmlElementBuilder
{
	/// <summary>
	/// Основной пакет документа OpenXML
	/// </summary>
	protected MainDocumentPart MainDocumentPart { get; }

	/// <summary>
	/// Тело документа
	/// </summary>
	protected Body Body => MainDocumentPart.Document.Body;
	
	/// <summary>
	/// Инициализирует <see cref="OpenXmlElementBuilder"/> по указанному <see cref="MainDocumentPart"/>
	/// </summary>
	/// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
	protected OpenXmlElementBuilder(MainDocumentPart mainDocumentPart)
	{
		MainDocumentPart = mainDocumentPart;
	}
}
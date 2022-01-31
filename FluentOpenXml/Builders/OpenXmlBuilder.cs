using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FluentOpenXml.Builders;

/// <summary>
/// Представляет абстрактный построитель элементов Open XML
/// </summary>
internal abstract class OpenXmlBuilder
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
	/// Инициализирует <see cref="OpenXmlBuilder"/> по указанному <see cref="MainDocumentPart"/>
	/// </summary>
	/// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
	protected OpenXmlBuilder(MainDocumentPart mainDocumentPart)
	{
		MainDocumentPart = mainDocumentPart;
	}
}
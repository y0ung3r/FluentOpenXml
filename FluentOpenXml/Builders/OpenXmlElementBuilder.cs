using DocumentFormat.OpenXml;
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

	/// <summary>
	/// Фабричный метод для создания указанного построителя
	/// </summary>
	/// <param name="elements">Элементы</param>
	/// <typeparam name="TBuilder">Тип построителя элемента</typeparam>
	private TBuilder CreateBuilder<TBuilder>(params OpenXmlElement[] elements)
		where TBuilder : OpenXmlElementBuilder
	{
		ArgumentNullException.ThrowIfNull(elements);
		
		return (TBuilder)Activator.CreateInstance
		(
			typeof(TBuilder), 
			MainDocumentPart, 
			elements
		);
	}

	/// <summary>
	/// Настраивает указанный элемент
	/// </summary>
	/// <param name="elements">Элементы</param>
	/// <param name="configure">Настраивает элемент</param>
	/// <typeparam name="TBuilder">Тип построителя элемента</typeparam>
	protected void ConfigureWith<TBuilder>(Action<TBuilder> configure, params OpenXmlElement[] elements)
		where TBuilder : OpenXmlElementBuilder
	{
		ArgumentNullException.ThrowIfNull(configure);
		ArgumentNullException.ThrowIfNull(elements);

		var builder = CreateBuilder<TBuilder>(elements);
		configure(builder);
	}
}
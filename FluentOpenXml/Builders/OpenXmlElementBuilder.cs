using System.Linq.Expressions;
using System.Reflection;
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
	/// <param name="element">Элемент</param>
	/// <typeparam name="TBuilder">Тип построителя элемента</typeparam>
	/// <typeparam name="TElement">Тип элемента</typeparam>
	private TBuilder CreateBuilder<TBuilder, TElement>(TElement element)
	{
		ArgumentNullException.ThrowIfNull(element);
		
		return (TBuilder)Activator.CreateInstance(typeof(TBuilder), MainDocumentPart, element);
	}

	/// <summary>
	/// Настраивает указанный элемент
	/// </summary>
	/// <param name="element">Элемент</param>
	/// <param name="configure">Настраивает элемент</param>
	/// <typeparam name="TBuilder">Тип построителя элемента</typeparam>
	/// <typeparam name="TElement">Тип элемента</typeparam>
	protected void Configure<TBuilder, TElement>(TElement element, Action<TBuilder> configure)
		where TElement : OpenXmlElement
	{
		ArgumentNullException.ThrowIfNull(element);
		ArgumentNullException.ThrowIfNull(configure);

		var builder = CreateBuilder<TBuilder, TElement>(element);
		configure(builder);
	}

	/// <summary>
	/// Устанавливает свойству элемента указанное значение, используя преобразование
	/// </summary>
	/// <param name="element">Элемент</param>
	/// <param name="expression">Путь к свойству</param>
	/// <param name="value">Новое значение</param>
	/// <param name="convert">Способ преобразования <see cref="TValue"/> в <see cref="TProperty"/></param>
	/// <typeparam name="TElement">Тип элемента</typeparam>
	/// <typeparam name="TProperty">Тип свойства</typeparam>
	/// <typeparam name="TValue">Тип нового значения</typeparam>
	protected void SetPropertyValue<TElement, TProperty, TValue>(TElement element, Expression<Func<TElement, TProperty>> expression, TValue value, Func<TValue, TProperty> convert)
		where TElement : OpenXmlElement
	{
		ArgumentNullException.ThrowIfNull(element);
		ArgumentNullException.ThrowIfNull(expression);
		ArgumentNullException.ThrowIfNull(value);
		
		var memberExpression = (MemberExpression)expression.Body;
		var propertyInfo = (PropertyInfo)memberExpression.Member;
        
		propertyInfo.SetValue
		(
			element, 
			convert(value)
		);
	}
}
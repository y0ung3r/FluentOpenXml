namespace FluentOpenXml.Builders.Interfaces;

/// <summary>
/// Представляет построитель документа
/// </summary>
public interface IDocumentBuilder
{
	/// <summary>
	/// Последняя секция в документе
	/// </summary>
	ISectionBuilder LastSection { get; }
	
	/// <summary>
	/// Вставляет разрыв секции в конец документа
	/// </summary>
	/// <returns></returns>
	ISectionBuilder AppendSectionBreak();
}
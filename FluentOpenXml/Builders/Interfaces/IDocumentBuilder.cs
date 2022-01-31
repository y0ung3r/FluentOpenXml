using FluentOpenXml.Builders.Section.Interfaces;

namespace FluentOpenXml.Builders.Interfaces;

/// <summary>
/// Представляет построитель документа
/// </summary>
public interface IDocumentBuilder
{
	/// <summary>
	/// Последняя секция в документе
	/// </summary>
	IDocumentBuilder GetLastSection(Action<ISectionBuilder> configureSection);

	/// <summary>
	/// Вставляет другой документ в конец текущего
	/// </summary>
	/// <param name="Stream">Последовательность байтов вставляемого документа</param>
	IDocumentBuilder AppendAnotherDocument(Stream stream);
	
	/// <summary>
	/// Вставляет разрыв секции в конец документа
	/// </summary>
	/// <returns></returns>
	IDocumentBuilder AppendSectionBreak(Action<ISectionBuilder> configureSection);
}
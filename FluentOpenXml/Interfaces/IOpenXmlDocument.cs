using FluentOpenXml.Builders.Interfaces;

namespace FluentOpenXml.Interfaces;

/// <summary>
/// Представляет документ
/// </summary>
public interface IOpenXmlDocument : IDisposable
{
	/// <summary>
	/// Возвращает построитель документа с помощью которого можно манипулировать документом
	/// </summary>
	/// <param name="edit">Действия по редактированию</param>
	IOpenXmlDocument Edit(Action<IDocumentBuilder> edit);
	
	/// <summary>
	/// Сохраняет документ в памяти
	/// </summary>
	IOpenXmlDocument Save();

	/// <summary>
	/// Закрывает документ
	/// </summary>
	void Close();
}
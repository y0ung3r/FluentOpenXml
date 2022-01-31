using FluentOpenXml.Builders.Interfaces;

namespace FluentOpenXml.Interfaces;

/// <summary>
/// Представляет документ
/// </summary>
public interface IOpenXmlDocument : IDisposable
{
	/// <summary>
	/// Открывает документ в памяти
	/// </summary>
	/// <param name="stream">Последовательность байтов в которой откроется документ</param>
	IOpenXmlDocument LoadFrom(Stream stream);

	/// <summary>
	/// Открывает документ из файла
	/// </summary>
	/// <param name="filepath">Путь к файлу</param>
	IOpenXmlDocument LoadFrom(string filepath);
	
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
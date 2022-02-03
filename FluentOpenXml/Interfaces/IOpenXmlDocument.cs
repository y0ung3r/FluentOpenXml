using FluentOpenXml.Builders.Interfaces;

namespace FluentOpenXml.Interfaces;

/// <summary>
/// Представляет документ
/// </summary>
public interface IOpenXmlDocument : IDisposable
{
	/// <summary>
	/// Настройки, применяемые к документам
	/// </summary>
	DocumentSettings Settings { get; }
	
	/// <summary>
	/// Указывает, что документ пуст 
	/// </summary>
	bool IsEmpty { get; }
	
	/// <summary>
	/// Возвращает построитель документа с помощью которого можно манипулировать документом
	/// </summary>
	/// <param name="edit">Метод, редактирующий документ</param>
	IOpenXmlDocument Edit(Action<IDocumentBuilder> edit);

	/// <summary>
	/// Сохраняет документ в указанный <see cref="Stream"/>, а затем открывает его в текущем контексте
	/// </summary>
	/// <param name="stream">Последовательность байтов</param>
	IOpenXmlDocument SaveTo(Stream stream);
	
	/// <summary>
	/// Сохраняет документ по указанному пути, а затем открывает его в текущем контексте
	/// </summary>
	/// <param name="path">Место сохранения в файловой системе</param>
	IOpenXmlDocument SaveTo(string path);
	
	/// <summary>
	/// Сохраняет документ 
	/// </summary>
	IOpenXmlDocument Save();

	/// <summary>
	/// Закрывает документ
	/// </summary>
	void Close();
}
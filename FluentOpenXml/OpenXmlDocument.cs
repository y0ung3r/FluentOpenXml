using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentOpenXml.Builders.Interfaces;
using FluentOpenXml.Interfaces;

namespace FluentOpenXml;

/// <summary>
/// Представляет стандартную реализацию <see cref="IOpenXmlDocument"/>
/// </summary>
public class OpenXmlDocument : IOpenXmlDocument
{
	/// <summary>
	/// Документ в памяти
	/// </summary>
	private Stream _stream;
	
	/// <summary>
	/// Документ
	/// </summary>
	private WordprocessingDocument _source;
	
	/// <summary>
	/// Стандартные настройки, применяемые к документу при открытии
	/// </summary>
	private OpenSettings DefaultOpenSettings { get; } = new()
	{
		AutoSave = false
	};

	/// <summary>
	/// Создает пустой документ
	/// </summary>
	public OpenXmlDocument()
	{
		_stream = new MemoryStream();
		
		_source = WordprocessingDocument.Create
		(
			_stream, 
			WordprocessingDocumentType.Document,
			DefaultOpenSettings.AutoSave
		);

		// TODO: загрузить стандартные стили в документ
	}

	/// <summary>
	/// Возвращает построитель документа с помощью которого можно манипулировать документом
	/// </summary>
	/// <param name="edit">Действия по редактированию</param>
	public IOpenXmlDocument Edit(Action<IDocumentBuilder> edit)
	{
		return this;
	}

	/// <summary>
	/// Сохраняет документ в памяти
	/// </summary>
	public IOpenXmlDocument Save()
	{
		throw new NotImplementedException();
	}
	
	/// <summary>
	/// Закрывает документ
	/// </summary>
	public void Close()
	{
		_source.Close();
	}

	/// <summary>
	/// Закрывает документ и выгружает все занятые им ресурсы
	/// </summary>
	public void Dispose()
	{
		_source.Dispose();
		_stream.Dispose();
	}
}
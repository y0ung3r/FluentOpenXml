using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Interfaces;
using FluentOpenXml.Exceptions;
using FluentOpenXml.Interfaces;

namespace FluentOpenXml;

/// <summary>
/// Представляет стандартную реализацию <see cref="IOpenXmlDocument"/>
/// </summary>
public class OpenXmlDocument : IOpenXmlDocument
{
	/// <summary>
	/// Стандартные настройки, применяемые к документам
	/// </summary>
	// TODO: отказаться от OpenSettings в пользу своего класса параметров
	private static OpenSettings DefaultSettings { get; } = new()
	{
		AutoSave = false
	};
	
	/// <summary>
	/// Указывает выгружены ли ресурсы документа
	/// </summary>
	private bool _isDisposed;

	/// <summary>
	/// Представление документа в виде последовательности байтов
	/// </summary>
	private Stream _stream;
	
	/// <summary>
	/// Определяет путь к файлу в файловой системе
	/// </summary>
	private string _filepath;
	
	/// <summary>
	/// Документ в формате OpenXML
	/// </summary>
	private WordprocessingDocument _source;

	/// <summary>
	/// Вставляет <see cref="MainDocumentPart"/> в документ
	/// </summary>
	private void AppendMainDocumentPart()
	{
		var mainDocumentPart = _source.AddMainDocumentPart();
		
		mainDocumentPart.Document = new Document();
	}

	/// <summary>
	/// Создает пустой документ в памяти
	/// </summary>
	public OpenXmlDocument()
		: this(
			new MemoryStream()
		)
	{ }
	
	/// <summary>
	/// Создает пустой документ в памяти, используя указанную последовательность байтов
	/// </summary>
	/// <param name="stream">Последовательность байтов в которой создастся документ</param>
	public OpenXmlDocument(Stream stream)
	{
		ArgumentNullException.ThrowIfNull(stream);

		_stream = stream;
		
		_source = WordprocessingDocument.Create
		(
			_stream, 
			WordprocessingDocumentType.Document,
			DefaultSettings.AutoSave
		);
		
		AppendMainDocumentPart();
		
		// TODO: загрузить стандартные стили в документ
	}

	/// <summary>
	/// Открывает документ в памяти
	/// </summary>
	/// <param name="stream">Последовательность байтов в которой откроется документ</param>
	public IOpenXmlDocument LoadFrom(Stream stream)
	{
		ArgumentNullException.ThrowIfNull(stream);
		
		_stream = stream;

		try
		{
			_source = WordprocessingDocument.Open
			(
				_stream, 
				isEditable: true,
				DefaultSettings
			);
		}
		
		catch (OpenXmlPackageException exception)
		{
			throw new InvalidDocumentException(exception.Message, exception);
		}

		return this;
	}

	/// <summary>
	/// Открывает документ из файла
	/// </summary>
	/// <param name="filepath">Путь к файлу</param>
	public IOpenXmlDocument LoadFrom(string filepath)
	{
		ArgumentNullException.ThrowIfNull(filepath);

		_filepath = filepath;

		var stream = new FileStream
		(
			_filepath,
			FileMode.Open,
			FileAccess.ReadWrite
		);

		return LoadFrom(stream);
	}

	/// <summary>
	/// Возвращает построитель документа с помощью которого можно манипулировать документом
	/// </summary>
	/// <param name="edit">Действия по редактированию</param>
	public IOpenXmlDocument Edit(Action<IDocumentBuilder> edit)
	{
		ArgumentNullException.ThrowIfNull(edit);

		ThrowIfDisposed();

		// TODO: прикрутить редактирование документа
		
		return this;
	}

	/// <summary>
	/// Сохраняет документ в памяти
	/// </summary>
	public IOpenXmlDocument Save()
	{
		ThrowIfDisposed();
		
		_source.Save();

		return this;
	}
	
	/// <summary>
	/// Закрывает документ
	/// </summary>
	public void Close()
	{
		ThrowIfDisposed();
		Dispose();
	}

	/// <summary>
	/// Бросает исключение <see cref="ObjectDisposedException"/>, если объект выгружен из памяти
	/// </summary>
	/// <exception cref="ObjectDisposedException">На момент выполнения действия объект был выгружен из памяти</exception>
	protected void ThrowIfDisposed()
	{
		if (_isDisposed)
		{
			throw new ObjectDisposedException
			(
				GetType().Name
			);
		}
	}

	/// <summary>
	/// Выгружает все занятые документом ресурсы
	/// </summary>
	public void Dispose()
	{
		if (!_isDisposed)
		{
			_source.Close();
			
			_stream.Dispose();
			_filepath = null;
			
			GC.SuppressFinalize(this);
		}
		
		_isDisposed = true;
	}
}
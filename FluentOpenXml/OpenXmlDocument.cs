// ReSharper disable MemberCanBePrivate.Global

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders;
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
    private DocumentSettings _settings;

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
    /// Основной пакет документа OpenXML
    /// </summary>
    private MainDocumentPart MainDocumentPart => _source.MainDocumentPart;

    /// <summary>
    /// Проверяет пуст ли документ 
    /// </summary>
    public bool IsEmpty
    {
        get
        {
            ThrowIfDisposed();
            EnsureMainDocumentPartAdded();

            return MainDocumentPart.Document.Body!.LastChild is null;
        }
    }

    /// <summary>
    /// Вставляет <see cref="MainDocumentPart"/> в документ
    /// </summary>
    private void AddMainDocumentPart()
    {
        var mainDocumentPart = _source.AddMainDocumentPart();

        mainDocumentPart.Document = new Document
        {
            Body = new Body()
        };
    }

    /// <summary>
    /// Проверяет добавлен ли основной пакет <see cref="MainDocumentPart"/> в документ и, в случае отсутствия, добавляет его
    /// </summary>
    private void EnsureMainDocumentPartAdded()
    {
        if (MainDocumentPart is null)
        {
            AddMainDocumentPart();
        }
    }

    /// <summary>
    /// Создает пустой документ, используя указанный <see cref="Stream"/>, а также применяет указанные настройки
    /// </summary>
    /// <param name="stream">Последовательность байтов в которой создастся документ</param>
    /// <param name="settings">Определяет настройки, применяемые к документу</param>
    private void Create(Stream stream, DocumentSettings settings)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(settings);

        _settings = settings;
        _stream = stream;

        _source = WordprocessingDocument.Create
        (
            _stream,
            WordprocessingDocumentType.Document,
            _settings.AllowAutoSaving
        );

        AddMainDocumentPart();

        // TODO: загрузить стандартные стили в документ
    }

    /// <summary>
    /// Создает пустой документ
    /// </summary>
    public OpenXmlDocument() => Create
    (
        new MemoryStream(),
        DocumentSettings.Default
    );

    /// <summary>
    /// Создает пустой документ и применяет указанные настройки
    /// </summary>
    /// <param name="settings">Определяет настройки, применяемые к документу</param>
    public OpenXmlDocument(DocumentSettings settings) => Create
    (
        new MemoryStream(),
        settings
    );

    /// <summary>
    /// Открывает документ
    /// </summary>
    /// <param name="stream">Последовательность байтов в которой откроется документ</param>
    public OpenXmlDocument(Stream stream) => LoadFrom(stream);

    /// <summary>
    /// Открывает документ и применяет указанные настройки
    /// </summary>
    /// <param name="stream">Последовательность байтов в которой откроется документ</param>
    /// <param name="settings">Определяет настройки, применяемые к документу</param>
    public OpenXmlDocument(Stream stream, DocumentSettings settings) => LoadFrom
    (
        stream,
        settings
    );

    /// <summary>
    /// Открывает документ
    /// </summary>
    /// <param name="stream">Последовательность байтов в которой откроется документ</param>
    public IOpenXmlDocument LoadFrom(Stream stream)
    {
        return LoadFrom
        (
            stream,
            DocumentSettings.Default
        );
    }

    /// <summary>
    /// Открывает документ и применяет указанные настройки
    /// </summary>
    /// <param name="stream">Последовательность байтов в которой откроется документ</param>
    /// <param name="settings">Определяет настройки, применяемые к документу</param>
    public IOpenXmlDocument LoadFrom(Stream stream, DocumentSettings settings)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(settings);

        _settings = settings;
        _stream = stream;

        try
        {
            _source = WordprocessingDocument.Open
            (
                _stream,
                isEditable: !_settings.IsReadOnly,
                new OpenSettings()
                {
                    AutoSave = _settings.AllowAutoSaving
                }
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
    public IOpenXmlDocument LoadFrom(string filepath) => LoadFrom
    (
        filepath,
        DocumentSettings.Default
    );

    /// <summary>
    /// Открывает документ из файла и применяет настройки
    /// </summary>
    /// <param name="filepath">Путь к файлу</param>
    /// <param name="settings">Определяет настройки, применяемые к документу</param>
    public IOpenXmlDocument LoadFrom(string filepath, DocumentSettings settings)
    {
        ArgumentNullException.ThrowIfNull(filepath);

        _filepath = filepath;

        var bytes = File.ReadAllBytes(_filepath);
        var stream = new MemoryStream(bytes);

        return LoadFrom(stream, settings);
    }

    /// <summary>
    /// Возвращает построитель документа с помощью которого можно манипулировать документом
    /// </summary>
    /// <param name="edit">Метод, редактирующий документ</param>
    public IOpenXmlDocument Edit(Action<IDocumentBuilder> edit)
    {
        ArgumentNullException.ThrowIfNull(edit);

        ThrowIfDisposed();
        ThrowIfReadOnly();
        EnsureMainDocumentPartAdded();

        // TODO: подключить внедрение зависимостей
        edit
        (
            new DocumentBuilder(MainDocumentPart)
        );

        return this;
    }

    /// <summary>
    /// Сохраняет документ в указанный <see cref="Stream"/>, а затем открывает его в текущем контексте
    /// </summary>
    /// <param name="stream">Последовательность байтов</param>
    public IOpenXmlDocument SaveTo(Stream stream)
    {
        ThrowIfDisposed();
        ThrowIfReadOnly();

        _source.Clone(stream);

        return LoadFrom
        (
            stream,
            _settings
        );
    }
    
    /// <summary>
    /// Сохраняет документ по указанному пути, а затем открывает его в текущем контексте
    /// </summary>
    /// <param name="path">Место сохранения в файловой системе</param>
    public IOpenXmlDocument SaveTo(string path) => SaveTo
    (
        // TODO: Я думаю, что это не лучшее решение. Подумать как по другому 
        new FileStream
        (
            path,
            FileMode.OpenOrCreate,
            FileAccess.ReadWrite
        )
    );
    
    /// <summary>
    /// Сохраняет текущий документ
    /// </summary>
    public IOpenXmlDocument Save()
    {
        // ReSharper disable once ConvertIfStatementToReturnStatement
        if (_filepath is not null)
        {
            return SaveTo(_filepath);
        }

        return SaveTo(_stream);
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
    /// Бросает исключение <see cref="DocumentInReadOnlyModeException"/>, если документ открыт в режиме «только для чтения»
    /// </summary>
    /// <exception cref="DocumentInReadOnlyModeException">Документ открыт в режиме «только для чтения»</exception>
    private void ThrowIfReadOnly()
    {
        if (_settings.IsReadOnly || !_stream.CanWrite)
        {
            throw new DocumentInReadOnlyModeException
            (
                "Вы не можете редактировать или сохранять документ в режиме «только для чтения»"
            );
        }
    }

    /// <summary>
    /// Бросает исключение <see cref="ObjectDisposedException"/>, если объект выгружен из памяти
    /// </summary>
    /// <exception cref="ObjectDisposedException">На момент выполнения действия объект был выгружен из памяти</exception>
    private void ThrowIfDisposed()
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
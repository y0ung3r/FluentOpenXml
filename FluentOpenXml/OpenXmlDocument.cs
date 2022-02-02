// ReSharper disable MemberCanBePrivate.Global

using System.IO.Packaging;
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
    /// Поле для <see cref="DocumentSettings" />
    /// </summary>
    private DocumentSettings _settings;

    /// <summary>
    /// Указывает выгружены ли ресурсы документа
    /// </summary>
    private bool _isDisposed;

    /// <summary>
    /// Представление документа в виде пакета данных
    /// </summary>
    private Package _package;

    /// <summary>
    /// Документ в формате OpenXML
    /// </summary>
    private WordprocessingDocument _source;

    /// <summary>
    /// Основной пакет документа OpenXML
    /// </summary>
    private MainDocumentPart MainDocumentPart => _source.MainDocumentPart;

    /// <inheritdoc cref="IOpenXmlDocument.IsEmpty"/>
    public bool IsEmpty
    {
        get
        {
            ThrowIfDisposed();
            EnsureMainDocumentPartAdded();

            return MainDocumentPart.Document.Body!.LastChild is null;
        }
    }

    /// <inheritdoc cref="IOpenXmlDocument.Settings"/>
    public DocumentSettings Settings => _settings;

    /// <summary>
    /// Создает пустой документ, используя указанный <see cref="Stream"/>, а также применяет указанные настройки
    /// </summary>
    /// <param name="stream">Последовательность байтов в которой создастся документ</param>
    /// <param name="settings">Определяет настройки, применяемые к документу</param>
    private void Create(Stream stream, DocumentSettings settings)
    {
        RememberOrThrowIfNull(ref _settings, ref settings);

        _source = WordprocessingDocument.Create
        (
            stream,
            WordprocessingDocumentType.Document,
            Settings.AllowAutoSaving
        );

        AddMainDocumentPart();

        // TODO: загрузить стандартные стили в документ
    }

    /// <summary>
    /// Создает пустой документ
    /// </summary>
    public OpenXmlDocument()
        : this(
            DocumentSettings.Default
        )
    { }

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
    public OpenXmlDocument(Stream stream)
        : this(
            stream,
            DocumentSettings.Default
        )
    { }

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
    /// Открывает документ по указанному пути
    /// </summary>
    /// <param name="filepath">Путь к файлу</param>
    public OpenXmlDocument(string filepath)
        : this(
            filepath, 
            DocumentSettings.Default
        )
    { }

    /// <summary>
    /// Открывает документ по указанному пути и применяет настройки
    /// </summary>
    /// <param name="filepath">Путь к документу</param>
    /// <param name="settings">Определяет настройки, применяемые к документу</param>
    public OpenXmlDocument(string filepath, DocumentSettings settings) => LoadFrom
    (
        filepath, 
        settings
    );

    /// <inheritdoc />
    public IOpenXmlDocument LoadFrom(Stream stream) => LoadFrom
    (
        stream,
        DocumentSettings.Default
    );

    /// <inheritdoc />
    public IOpenXmlDocument LoadFrom(Stream stream, DocumentSettings settings)
    {
        RememberOrThrowIfNull(ref _settings, ref settings);

        _package = Package.Open
        (
            stream,
            Settings.DocumentMode,
            Settings.DocumentAccess
        );

        try
        {
            _source = WordprocessingDocument.Open
            (
                _package,
                new OpenSettings()
                {
                    AutoSave = Settings.AllowAutoSaving
                }
            );
        }

        catch (OpenXmlPackageException exception)
        {
            throw new InvalidDocumentException(exception.Message, exception);
        }

        return this;
    }

    /// <inheritdoc />
    public IOpenXmlDocument LoadFrom(string filepath) => LoadFrom
    (
        filepath,
        DocumentSettings.Default
    );

    /// <inheritdoc />
    public IOpenXmlDocument LoadFrom(string filepath, DocumentSettings settings)
    {
        var bytes = File.ReadAllBytes(filepath);
        var stream = new MemoryStream(bytes);

        return LoadFrom(stream, settings);
    }

    /// <inheritdoc />
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
    /// Сохраняет документ в указанный пакет данных
    /// </summary>
    /// <param name="package">Пакет данных</param>
    private void SaveTo(Package package)
    {
        ArgumentNullException.ThrowIfNull(package);
        
        ThrowIfDisposed();
        ThrowIfReadOnly();

        _source.Clone(package);
    }

    /// <inheritdoc />
    public IOpenXmlDocument SaveTo(Stream stream)
    {
        SaveTo
        (
            Package.Open
            (
                stream,
                Settings.DocumentMode,
                Settings.DocumentAccess
            )
        );

        return LoadFrom
        (
            stream,
            _settings
        );
    }

    /// <inheritdoc />
    public IOpenXmlDocument SaveTo(string path) => SaveTo
    (
        new FileStream
        (
            path,
            Settings.DocumentMode,
            Settings.DocumentAccess
        )
    );

    /// <inheritdoc />
    public IOpenXmlDocument Save()
    {
        SaveTo(_package);

        return this;
    }

    /// <inheritdoc />
    public void Close()
    {
        ThrowIfDisposed();
        Dispose();
    }

    /// <summary>
    /// Выгружает все занятые документом ресурсы
    /// </summary>
    public void Dispose()
    {
        if (!_isDisposed)
        {
            _source.Close(); // Закрывает и _package тоже
            _package = null;

            GC.SuppressFinalize(this);
        }

        _isDisposed = true;
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
    /// Сохраняет указанное значение в поле документа. Если значение null, выбрасывает исключение <see cref="ArgumentNullException"/>
    /// </summary>
    /// <param name="field">Поле</param>
    /// <param name="value">Значение</param>
    /// <typeparam name="TField">Тип поля</typeparam>
    // ReSharper disable once RedundantAssignment
    private void RememberOrThrowIfNull<TField>(ref TField field, ref TField value)
    {
        ArgumentNullException.ThrowIfNull(value);

        field = value;
    }

    /// <summary>
    /// Бросает исключение <see cref="DocumentInReadOnlyModeException"/>, если документ открыт в режиме «только для чтения»
    /// </summary>
    /// <exception cref="DocumentInReadOnlyModeException">Документ открыт в режиме «только для чтения»</exception>
    private void ThrowIfReadOnly()
    {
        if (Settings.IsReadOnly || _package.FileOpenAccess.Equals(FileAccess.Read))
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
}
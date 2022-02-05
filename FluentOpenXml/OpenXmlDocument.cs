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
public sealed class OpenXmlDocument : IOpenXmlDocument
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
    /// Документ в формате OpenXML
    /// </summary>
    private WordprocessingDocument _source;

    /// <inheritdoc cref="IOpenXmlDocument.Settings"/>
    public DocumentSettings Settings => _settings;

    /// <inheritdoc cref="IOpenXmlDocument.IsEmpty"/>
    public bool IsEmpty
    {
        get
        {
            ThrowIfDisposed();
            EnsureMainDocumentPartAdded();
            
            return _source.MainDocumentPart!.Document.Body!.LastChild is null;
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
    /// Открывает документ по указанному пакету данных, используя примененные настройки
    /// </summary>
    /// <param name="package">Пакет данных</param>
    private IOpenXmlDocument LoadFrom(Package package)
    {
        ArgumentNullException.ThrowIfNull(package);
        
        try
        {
            _source = WordprocessingDocument.Open
            (
                package,
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

    /// <summary>
    /// Открывает документ и применяет указанные настройки
    /// </summary>
    /// <param name="stream">Последовательность байтов в которой откроется документ</param>
    /// <param name="settings">Определяет настройки, применяемые к документу</param>
    private IOpenXmlDocument LoadFrom(Stream stream, DocumentSettings settings)
    {
        ArgumentNullException.ThrowIfNull(stream);
        RememberOrThrowIfNull(ref _settings, ref settings);

        return LoadFrom
        (
            Package.Open
            (
                stream,
                Settings.DocumentMode,
                Settings.DocumentAccess
            )
        );
    }

    /// <summary>
    /// Открывает документ по указанному пути и применяет настройки
    /// </summary>
    /// <param name="filepath">Путь к документу</param>
    /// <param name="settings">Определяет настройки, применяемые к документу</param>
    // ReSharper disable once UnusedMethodReturnValue.Local
    private IOpenXmlDocument LoadFrom(string filepath, DocumentSettings settings)
    {
        ArgumentNullException.ThrowIfNull(filepath);

        var bytes = File.ReadAllBytes(filepath);
        var stream = new MemoryStream(bytes);

        return LoadFrom(stream, settings);
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
    public IOpenXmlDocument Edit(Action<IDocumentBuilder> edit)
    {
        ArgumentNullException.ThrowIfNull(edit);

        ThrowIfDisposed();
        ThrowIfReadOnly();
        EnsureMainDocumentPartAdded();

        var documentBuilder = new DocumentBuilder(_source.MainDocumentPart);
        edit(documentBuilder);

        return this;
    }

    /// <summary>
    /// Сохраняет документ в указанный пакет данных
    /// </summary>
    /// <param name="package">Пакет данных</param>
    private IOpenXmlDocument SaveTo(Package package)
    {
        ArgumentNullException.ThrowIfNull(package);
        
        ThrowIfDisposed();
        ThrowIfReadOnly();

        _source.Clone(package);

        return LoadFrom(package);
    }

    /// <inheritdoc />
    public IOpenXmlDocument SaveTo(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

        var package = Package.Open
        (
            stream,
            Settings.DocumentMode,
            Settings.DocumentAccess
        );

        return SaveTo(package);
    }

    /// <inheritdoc />
    public IOpenXmlDocument SaveTo(string path)
    {
        ArgumentNullException.ThrowIfNull(path);

        var stream = new FileStream
        (
            path,
            Settings.DocumentMode,
            Settings.DocumentAccess
        );

        return SaveTo(stream);
    }

    /// <inheritdoc />
    public IOpenXmlDocument Save() => SaveTo
    (
        _source.Package
    );

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
            _source.Close();
            _source = null;
            _settings = null;

            // ReSharper disable once GCSuppressFinalizeForTypeWithoutDestructor
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
        if (_source.MainDocumentPart is null)
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
        if (Settings.IsReadOnly || _source.Package.FileOpenAccess.Equals(FileAccess.Read))
        {
            throw new DocumentInReadOnlyModeException
            (
                "В режиме «только для чтения» невозможно изменять состояние документа"
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
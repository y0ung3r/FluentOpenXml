// ReSharper disable MemberCanBePrivate.Global

namespace FluentOpenXml;

/// <summary>
/// Настройки для документа
/// </summary>
public class DocumentSettings
{
    /// <summary>
    /// Стандартные настройки
    /// </summary>
    public static DocumentSettings Default => new()
    {
        IsReadOnly = false,
        AllowAutoSaving = false,
        AllowUpdateFieldsOnOpen = true
    };

    /// <summary>
    /// Поле для <see cref="AllowAutoSaving"/>
    /// </summary>
    private readonly bool _allowAutoSaving;

    /// <summary>
    /// Способ открытия открытия для документа
    /// </summary>
    internal FileMode DocumentMode => IsReadOnly ? FileMode.Open : FileMode.OpenOrCreate;
    
    /// <summary>
    /// Предоставляемый доступ к файловой системе для документа
    /// </summary>
    internal FileAccess DocumentAccess => IsReadOnly ? FileAccess.Read : FileAccess.ReadWrite;
    
    /// <summary>
    /// Указывает, что документ должен открываться только для просмотра
    /// </summary>
    public bool IsReadOnly { get; init; }

    /// <summary>
    /// Разрешить автосохранение документа при закрытии и открытии документа.
    /// В режиме «только для чтения» эта настройка будет автоматически отключена 
    /// </summary>
    public bool AllowAutoSaving
    {
        get => _allowAutoSaving && !IsReadOnly;
        init => _allowAutoSaving = value;
    }

    /// <summary>
    /// Разрешить обновление полей в оглавлении при открытии документа
    /// </summary>
    public bool AllowUpdateFieldsOnOpen { get; init; }
}
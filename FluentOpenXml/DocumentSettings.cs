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
    /// Указывает, что документ должен открываться только для просмотра
    /// </summary>
    public bool IsReadOnly { get; init; }

    /// <summary>
    /// Разрешить автосохранение документа при закрытии и открытии документа
    /// </summary>
    public bool AllowAutoSaving { get; init; }

    /// <summary>
    /// Разрешить обновление полей в оглавлении при открытии документа
    /// </summary>
    public bool AllowUpdateFieldsOnOpen { get; init; }
}
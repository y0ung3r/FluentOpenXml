namespace FluentOpenXml.Exceptions;

/// <summary>
/// Представляет исключение, возникающее при попытке открыть не поддерживаемый форматом OpenXML документ
/// </summary>
public class InvalidDocumentException : Exception
{
    /// <summary>
    /// Инициализирует <see cref="InvalidDocumentException"/> с указанным сообщением
    /// </summary>
    /// <param name="message">Сообщение</param>
    public InvalidDocumentException(string message)
        : base(message)
    { }

    /// <summary>
    /// Инициализирует <see cref="InvalidDocumentException"/> с указанным сообщением и внутренним исключением
    /// </summary>
    /// <param name="message">Сообщение</param>
    /// <param name="innerException">Внутреннее исключение</param>
    public InvalidDocumentException(string message, Exception innerException)
        : base(message, innerException)
    { }
}
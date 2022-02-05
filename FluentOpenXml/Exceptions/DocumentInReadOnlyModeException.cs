namespace FluentOpenXml.Exceptions;

/// <summary>
/// Представляет исключение, возникающее при попытке редактировать документ в режиме только для чтения
/// </summary>
public sealed class DocumentInReadOnlyModeException : Exception
{
   /// <summary>
   /// Инициализирует <see cref="DocumentInReadOnlyModeException"/> с указанным сообщением
   /// </summary>
   /// <param name="message">Сообщение</param>
   public DocumentInReadOnlyModeException(string message)
      : base(message)
   { }

   /// <summary>
   /// Инициализирует <see cref="DocumentInReadOnlyModeException"/> с указанным сообщением и внутренним исключением
   /// </summary>
   /// <param name="message">Сообщение</param>
   /// <param name="innerException">Внутреннее исключение</param>
   public DocumentInReadOnlyModeException(string message, Exception innerException)
      : base(message, innerException)
   { }
}
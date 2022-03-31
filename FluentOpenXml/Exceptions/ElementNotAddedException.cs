namespace FluentOpenXml.Exceptions;

/// <summary>
/// Представляет исключение, возникающее при попытке редактировать элемент, который еще не добавлен
/// </summary>
public sealed class ElementNotAddedException : Exception
{
   /// <summary>
   /// Инициализирует <see cref="DocumentInReadOnlyModeException"/> с указанным сообщением и внутренним исключением
   /// </summary>
   /// <param name="elementName">Имя элемента</param>
   /// <param name="innerException">Внутреннее исключение</param>
   public ElementNotAddedException(string elementName, Exception innerException = default)
      : base($"Элемент {elementName} не добавлен в документ", innerException)
   { }
}
using DocumentFormat.OpenXml;

namespace FluentOpenXml.Extensions;

/// <summary>
/// Методы-расширения для <see cref="OpenXmlElement"/>
/// </summary>
internal static class OpenXmlElementExtensions
{
    /// <summary>
    /// Возвращает дочерние элементы указанного типа
    /// </summary>
    /// <param name="xmlElement">Родительский элемент</param>
    /// <typeparam name="TChild">Тип дочернего элемента</typeparam>
    internal static IEnumerable<TChild> GetChildren<TChild>(this OpenXmlElement xmlElement)
        where TChild : OpenXmlElement
    {
        return xmlElement.ChildElements.OfType<TChild>();
    }

    /// <summary>
    /// Возвращает первый дочерний элемент указанного типа или стандартное значение, если последовательность дочерних элементов пуста
    /// </summary>
    /// <param name="xmlElement">Родительский элемент</param>
    /// <typeparam name="TChild">Тип дочернего элемента</typeparam>
    internal static TChild FirstOrDefaultChild<TChild>(this OpenXmlElement xmlElement)
        where TChild : OpenXmlElement
    {
        return GetChildren<TChild>(xmlElement).FirstOrDefault();
    }

    /// <summary>
    /// Возвращает последний дочерний элемент указанного типа или стандартное значение, если последовательность дочерних элементов пуста
    /// </summary>
    /// <param name="xmlElement">Родительский элемент</param>
    /// <typeparam name="TChild">Тип дочернего элемента</typeparam>
    internal static TChild LastOrDefaultChild<TChild>(this OpenXmlElement xmlElement)
        where TChild : OpenXmlElement
    {
        return GetChildren<TChild>(xmlElement).LastOrDefault();
    }

    /// <summary>
    /// Возвращает дочерний элемент указанного типа используя заданный индекс в последовательности дочерних элементов
    /// </summary>
    /// <param name="xmlElement">Родительский элемент</param>
    /// <param name="index">Индекс</param>
    /// <typeparam name="TChild">Тип дочернего элемента</typeparam>
    internal static TChild ChildAt<TChild>(this OpenXmlElement xmlElement, int index)
        where TChild : OpenXmlElement
    {
        return GetChildren<TChild>(xmlElement).ElementAt(index);
    }

    /// <summary>
    /// Вставляет новый дочерний элемент в конец последовательности
    /// </summary>
    /// <param name="xmlElement">Родительский элемент</param>
    /// <typeparam name="TChild">Тип дочернего элемента</typeparam>
    internal static TChild AppendChild<TChild>(this OpenXmlElement xmlElement)
        where TChild : OpenXmlElement
    {
        var child = Activator.CreateInstance<TChild>();

        xmlElement.Append(child);

        return child;
    }

    /// <summary>
    /// Вставляет новый дочерний элемент перед указанным элементом последовательности
    /// </summary>
    /// <param name="xmlElement">Родительский элемент</param>
    /// <param name="relativeChild">Элемент перед которым необходимо вставить новый дочерний элемент</param>
    /// <typeparam name="TChild">Тип дочернего элемента</typeparam>
    internal static TChild PrependChild<TChild>(this OpenXmlElement xmlElement, OpenXmlElement relativeChild)
        where TChild : OpenXmlElement
    {
        var child = Activator.CreateInstance<TChild>();

        xmlElement.InsertBefore(child, relativeChild);

        return child;
    }

    /// <summary>
    /// Возвращает первый дочерний элемент указанного типа или вставляет новый элемент того же типа перед указанным элементом последовательности
    /// </summary>
    /// <param name="xmlElement">Родительский элемент</param>
    /// <param name="relativeChild">Элемент перед которым необходимо вставить новый дочерний элемент</param>
    /// <typeparam name="TChild">Тип дочернего элемента</typeparam>
    internal static TChild FirstOrPrependChild<TChild>(this OpenXmlElement xmlElement, OpenXmlElement relativeChild)
        where TChild : OpenXmlElement
    {
        return FirstOrDefaultChild<TChild>(xmlElement) ?? PrependChild<TChild>(xmlElement, relativeChild);
    }

    /// <summary>
    /// Клонирует дочерний элемент и вставляет его в конец последовательности
    /// </summary>
    /// <param name="xmlElement">Родительский элемент</param>
    /// <param name="child">Дочерний элемент</param>
    /// <typeparam name="TChild">Тип дочернего элемента</typeparam>
    internal static TChild CloneAndAppendChild<TChild>(this OpenXmlElement xmlElement, OpenXmlElement child)
        where TChild : OpenXmlElement
    {
        var clone = CloneAs<TChild>(child);
        return xmlElement.AppendChild(clone);
    }

    /// <summary>
    /// Возвращает элемент указанного типа или вставляет новый элемент в конец последовательности
    /// </summary>
    /// <param name="xmlElement">Родительский элемент</param>
    /// <typeparam name="TChild">Тип дочернего элемента</typeparam>
    internal static TChild FirstOrNewChild<TChild>(this OpenXmlElement xmlElement)
        where TChild : OpenXmlElement
    {
        return FirstOrDefaultChild<TChild>(xmlElement) ?? AppendChild<TChild>(xmlElement);
    }

    /// <summary>
    /// Возвращает родительские элементы указанного типа
    /// </summary>
    /// <param name="xmlElement">Дочерний элемент</param>
    /// <typeparam name="TAncestor">Тип родительского элемента</typeparam>
    internal static IEnumerable<TAncestor> GetAncestors<TAncestor>(this OpenXmlElement xmlElement)
        where TAncestor : OpenXmlElement
    {
        return xmlElement.Ancestors<TAncestor>();
    }

    /// <summary>
    /// Возвращает первый родительский элемент указанного типа или стандартное значение, если последовательность родительских элементов пуста
    /// </summary>
    /// <param name="xmlElement">Дочерний элемент</param>
    /// <typeparam name="TAncestor">Тип родительского элемента</typeparam>
    internal static TAncestor FirstOrDefaultAncestor<TAncestor>(this OpenXmlElement xmlElement)
        where TAncestor : OpenXmlElement
    {
        return GetAncestors<TAncestor>(xmlElement).FirstOrDefault();
    }

    /// <summary>
    /// Возвращает последний родительский элемент указанного типа или стандартное значение, если последовательность родительских элементов пуста
    /// </summary>
    /// <param name="xmlElement">Дочерний элемент</param>
    /// <typeparam name="TAncestor">Тип родительского элемента</typeparam>
    internal static TAncestor LastOrDefaultAncestor<TAncestor>(this OpenXmlElement xmlElement)
        where TAncestor : OpenXmlElement
    {
        return GetAncestors<TAncestor>(xmlElement).LastOrDefault();
    }

    /// <summary>
    /// Клонирует элемент указанного типа
    /// </summary>
    /// <param name="xmlElement">Элемент</param>
    /// <typeparam name="TXmlElement">Тип элемента</typeparam>
    internal static TXmlElement CloneAs<TXmlElement>(this OpenXmlElement xmlElement)
        where TXmlElement : OpenXmlElement
    {
        return xmlElement.Clone() as TXmlElement;
    }
}
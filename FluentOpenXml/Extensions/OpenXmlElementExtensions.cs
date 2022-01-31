using DocumentFormat.OpenXml;

namespace FluentOpenXml.Extensions;

/// <summary>
/// Методы-расширения для <see cref="OpenXmlElement"/>
/// </summary>
internal static class OpenXmlElementExtensions
{
    // TODO: добавить комментарии
    
    internal static IEnumerable<TChild> GetChildren<TChild>(this OpenXmlElement xmlElement)
        where TChild : OpenXmlElement
    {
        return xmlElement.ChildElements.OfType<TChild>();
    }

    internal static TChild FirstOrDefaultChild<TChild>(this OpenXmlElement xmlElement)
        where TChild : OpenXmlElement
    {
        return GetChildren<TChild>(xmlElement).FirstOrDefault();
    }

    internal static TChild LastOrDefaultChild<TChild>(this OpenXmlElement xmlElement)
        where TChild : OpenXmlElement
    {
        return GetChildren<TChild>(xmlElement).LastOrDefault();
    }

    internal static TChild ChildAt<TChild>(this OpenXmlElement xmlElement, int index)
        where TChild : OpenXmlElement
    {
        return GetChildren<TChild>(xmlElement).ElementAt(index);
    }

    internal static TChild AppendChild<TChild>(this OpenXmlElement xmlElement)
        where TChild : OpenXmlElement
    {
        var child = Activator.CreateInstance<TChild>();

        xmlElement.Append(child);

        return child;
    }

    internal static TChild PrependChild<TChild>(this OpenXmlElement xmlElement, OpenXmlElement relativeChild)
        where TChild : OpenXmlElement
    {
        var child = Activator.CreateInstance<TChild>();

        xmlElement.InsertBefore(child, relativeChild);

        return child;
    }

    internal static TChild FirstOrPrependChild<TChild>(this OpenXmlElement xmlElement, OpenXmlElement relativeChild)
        where TChild : OpenXmlElement
    {
        return FirstOrDefaultChild<TChild>(xmlElement) ?? PrependChild<TChild>(xmlElement, relativeChild);
    }

    internal static TChild CloneAndAppendChild<TChild>(this OpenXmlElement xmlElement, TChild child)
        where TChild : OpenXmlElement
    {
        var clone = CloneAs<TChild>(child);
        return xmlElement.AppendChild(clone);
    }

    internal static TChild FirstOrNewChild<TChild>(this OpenXmlElement xmlElement)
        where TChild : OpenXmlElement
    {
        return FirstOrDefaultChild<TChild>(xmlElement) ?? AppendChild<TChild>(xmlElement);
    }

    internal static IEnumerable<TAncestor> GetAncestors<TAncestor>(this OpenXmlElement xmlElement)
        where TAncestor : OpenXmlElement
    {
        return xmlElement.Ancestors<TAncestor>();
    }

    internal static TAncestor FirstOrDefaultAncestor<TAncestor>(this OpenXmlElement xmlElement)
        where TAncestor : OpenXmlElement
    {
        return GetAncestors<TAncestor>(xmlElement).FirstOrDefault();
    }

    internal static TAncestor LastOrDefaultAncestor<TAncestor>(this OpenXmlElement xmlElement)
        where TAncestor : OpenXmlElement
    {
        return GetAncestors<TAncestor>(xmlElement).LastOrDefault();
    }

    internal static TXmlElement CloneAs<TXmlElement>(this OpenXmlElement xmlElement)
        where TXmlElement : OpenXmlElement
    {
        return xmlElement.Clone() as TXmlElement;
    }
}
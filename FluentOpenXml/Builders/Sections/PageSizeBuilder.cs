﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Sections.Interfaces;
using FluentOpenXml.Units.Extensions;
using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.Sections;

/// <summary>
/// Стандартная реализация <see cref="IPageSizeBuilder"/>
/// </summary>
internal class PageSizeBuilder : OpenXmlElementBuilder, IPageSizeBuilder
{
    /// <summary>
    /// Размеры страницы
    /// </summary>
    private readonly PageSize _pageSize;

    /// <summary>
    /// Ширина страницы в <see cref="Twips"/>
    /// </summary>
    private Twips Width
    {
        get => new Twips(_pageSize.Width);
        set => _pageSize.Width = value;
    }
    
    /// <summary>
    /// Высота страницы в <see cref="Twips"/>
    /// </summary>
    private Twips Height
    {
        get => new Twips(_pageSize.Height);
        set => _pageSize.Height = value;
    }

    /// <summary>
    /// Инициализирует <see cref="PageSizeBuilder"/> по указанному <see cref="MainDocumentPart"/> и <see cref="PageSize"/>
    /// </summary>
    /// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
    /// <param name="pageSize">Размеры страницы</param>
    public PageSizeBuilder(MainDocumentPart mainDocumentPart, PageSize pageSize) 
        : base(mainDocumentPart)
    {
        _pageSize = pageSize;
    }

    /// <inheritdoc />
    public TUnits GetWidth<TUnits>()
        where TUnits : UniversalUnits => Width.ToEmu().To<TUnits>();

    /// <inheritdoc />
    public IPageSizeBuilder SetWidth<TUnits>(double value)
        where TUnits : UniversalUnits
    {
        Width = value.As<TUnits>().ToEmu().ToTwips();

        return this;
    }

    /// <inheritdoc />
    public TUnits GetHeight<TUnits>() 
        where TUnits : UniversalUnits => Height.ToEmu().To<TUnits>();

    /// <inheritdoc />
    public IPageSizeBuilder SetHeight<TUnits>(double value)
        where TUnits : UniversalUnits
    {
        Height = value.As<TUnits>().ToEmu().ToTwips();

        return this;
    }
}
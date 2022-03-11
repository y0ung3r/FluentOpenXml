﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentOpenXml.Builders.Paragraphs.Interfaces;
using FluentOpenXml.Builders.Sections.Enums;
using FluentOpenXml.Builders.Sections.Interfaces;
using FluentOpenXml.Extensions;
using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Builders.Sections;

/// <summary>
/// Стандартная реализация <see cref="ISectionBuilder"/>
/// </summary>
internal class SectionBuilder : OpenXmlElementBuilder, ISectionBuilder
{
	/// <summary>
	/// Настраиваемая секция
	/// </summary>
	private readonly SectionProperties _section;

	/// <summary>
	/// Последний абзац в документе
	/// </summary>
	private Paragraph LastParagraph => Body.LastOrDefaultChild<Paragraph>();

	/// <summary>
	/// Инициализирует <see cref="SectionBuilder"/> по указанному <see cref="MainDocumentPart"/>
	/// </summary>
	/// <param name="mainDocumentPart">Основной пакет документа OpenXML</param>
	/// <param name="section">Настраиваемая секция</param>
	internal SectionBuilder(MainDocumentPart mainDocumentPart, SectionProperties section)
		: base(mainDocumentPart)
	{
		_section = section;
	}
	
	/// <inheritdoc/>
	public ISectionBuilder SetPageOrientation(PageOrientation orientation)
	{
		SetPageSize(builder =>
		{
			var width = builder.GetWidth<Twips>();
			var height = builder.GetHeight<Twips>();

			builder.SetWidth<Twips>(height.Value)
				   .SetHeight<Twips>(width.Value);
		});
		
		return this;
	}
	
	/// <inheritdoc/>
	public ISectionBuilder SetPageSize(Action<IPageSizeBuilder> setPageSize)
	{
		var pageSize = _section.FirstOrNewChild<PageSize>();
		Configure<PageSizeBuilder, PageSize>(pageSize, setPageSize);

		return this;
	}

	/// <inheritdoc/>
	public ISectionBuilder SetPageMargin(Action<IPageMarginBuilder> setPageMargin)
	{
		var pageMargin = _section.FirstOrNewChild<PageMargin>();
		Configure<PageMarginBuilder, PageMargin>(pageMargin, setPageMargin);

		return this;
	}

	/// <inheritdoc/>
	public ISectionBuilder ConfigureLastParagraph(Action<IParagraphBuilder> configureParagraph)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc/>
	public ISectionBuilder AppendParagraph(Action<IParagraphBuilder> configureParagraph)
	{
		throw new NotImplementedException();
	}

	/// <inheritdoc/>
	public ISectionBuilder AppendChunk(Action<string> configureRelationship)
	{
		var identificator = Guid.NewGuid().ToString();
		configureRelationship(identificator);
		
		var chunk = new AltChunk()
		{
			Id = identificator
		};

		Body.InsertAfter(chunk, LastParagraph);

		return this;
	}
}
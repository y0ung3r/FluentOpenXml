using FluentOpenXml.Builders.Sections.Enums;

namespace FluentOpenXml.Builders.PageLayout.Interfaces;

/// <summary>
/// Представляет построитель страницы
/// </summary>
public interface IPageLayoutBuilder : ITransposable
{
	/// <summary>
	/// Устанавливает указанную ориентацию для страницы
	/// </summary>
	/// <param name="orientation">Ориентация</param>
	IPageLayoutBuilder SetOrientation(PageOrientation orientation);
	
	/// <summary>
	/// Устанавливает размеры страницы в соответствии с заданными параметры
	/// </summary>
	/// <param name="setSize">Настраивает размеры страницы</param>
	IPageLayoutBuilder SetSize(Action<IPageSizeBuilder> setSize);

	/// <summary>
	/// Устанавливает внешние отступы у страницы в соответствии с заданными параметрами
	/// </summary>
	/// <param name="setMargin">Настраивает внешние границы</param>
	IPageLayoutBuilder SetMargin(Action<IPageMarginBuilder> setMargin);
}
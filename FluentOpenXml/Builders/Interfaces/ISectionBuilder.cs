namespace FluentOpenXml.Builders.Interfaces;

/// <summary>
/// Представляет построитель секции в документе
/// </summary>
public interface ISectionBuilder
{
	/// <summary>
	/// Устанавливает ориентацию страницы
	/// </summary>
	ISectionBuilder SetPageOrientation();
}
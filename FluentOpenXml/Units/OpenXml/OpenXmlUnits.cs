namespace FluentOpenXml.Units.OpenXml;

/// <summary>
/// Представляет базовый класс для единиц измерения, используемых в Open XML
/// </summary>
internal abstract class OpenXmlUnits
{
	/// <summary>
	/// Значение
	/// </summary>
	internal long Value { get; }

	/// <summary>
	/// Инициализирует <see cref="OpenXmlUnits"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal OpenXmlUnits(double value)
	{
		if (value < 0)
		{
			throw new ArgumentException($"Значение \"{nameof(value)}\" не может быть меньше нуля");
		}
		
		Value = Convert.ToInt64
		(
			Math.Truncate(value)
		);
	}
}
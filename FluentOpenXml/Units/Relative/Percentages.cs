namespace FluentOpenXml.Units.Relative;

/// <summary>
/// Представляет проценты
/// </summary>
internal sealed class Percentages : FloatingPointUnits
{
	/// <summary>
	/// Инициализирует <see cref="Percentages"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Percentages(double value)
		: base(value)
	{ }

	/// <summary>
	/// Преобразует <see cref="Percentages"/> в <see cref="FiftiethsOfAPercentage"/>
	/// </summary>
	internal FiftiethsOfAPercentage ToFiftiethsOfAPercentage() => new FiftiethsOfAPercentage(Value * 100.0 / 2.0);
}
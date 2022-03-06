namespace FluentOpenXml.Units.Relative;

/// <summary>
/// Представляет пятидесятые доли процента
/// </summary>
internal sealed class FiftiethsOfAPercentage : FloatingPointUnits
{
	/// <summary>
	/// Инициализирует <see cref="FiftiethsOfAPercentage"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal FiftiethsOfAPercentage(double value)
		: base(value)
	{ }

	/// <summary>
	/// Преобразует <see cref="FiftiethsOfAPercentage"/> в <see cref="Percentages"/>
	/// </summary>
	internal Percentages ToPercentages() => new Percentages(Value * 2.0 / 100.0);
}
namespace FluentOpenXml.Units;

/// <summary>
/// Представляет проценты
/// </summary>
internal readonly struct Percentages
{
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="Percentages"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Percentages(double value)
	{
		Value = value;
	}

	/// <summary>
	/// Преобразует <see cref="Percentages"/> в <see cref="Pcts"/>
	/// </summary>
	internal Pcts ToPcts() => new Pcts(Value / 0.02);
}
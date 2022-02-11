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
		if (value < 0)
		{
			throw new ArgumentException($"Значение \"{nameof(value)}\" не может быть меньше нуля");
		}
		
		Value = value;
	}

	/// <summary>
	/// Преобразует <see cref="Percentages"/> в <see cref="FiftiethsOfAPercent"/>
	/// </summary>
	internal FiftiethsOfAPercent ToFiftiethsOfAPercent() => new FiftiethsOfAPercent(Value / 0.02);
}
namespace FluentOpenXml.Units;

/// <summary>
/// Представляет пятидесятые доли процента (пункты)
/// Используется для относительных измерений: ширина таблицы, ячейки, поля страницы и т. д.
/// </summary>
internal readonly struct Pcts
{
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="Pcts"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Pcts(double value)
	{
		Value = value;
	}

	/// <summary>
	/// Преобразует <see cref="Pcts"/> в <see cref="Percentages"/>
	/// </summary>
	internal Percentages ToPercentages() => new Percentages(Value * 0.02);
}
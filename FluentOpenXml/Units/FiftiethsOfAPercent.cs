namespace FluentOpenXml.Units;

/// <summary>
/// Представляет пятидесятые доли процента
/// Используется для относительных измерений: ширина таблицы, ячейки, поля страницы и т. д.
/// </summary>
internal readonly struct FiftiethsOfAPercent
{
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="FiftiethsOfAPercent"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal FiftiethsOfAPercent(double value)
	{
		Value = value;
	}
}
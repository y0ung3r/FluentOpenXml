namespace FluentOpenXml.Units;

/// <summary>
/// Представляет половины от точки
/// Используется в размерах шрифта
/// </summary>
internal readonly struct HalfPoints
{
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="HalfPoints"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal HalfPoints(double value)
	{
		Value = value;
	}
	
	/// <summary>
	/// Преобразует <see cref="HalfPoints"/> в <see cref="Points"/>
	/// </summary>
	internal Points ToPoints() => new Points(Value / 2.0);
}
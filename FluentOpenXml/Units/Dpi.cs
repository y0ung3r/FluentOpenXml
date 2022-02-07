namespace FluentOpenXml.Units;

/// <summary>
/// Представляет количество точек на дюйм
/// </summary>
internal readonly struct Dpi
{
	/// <summary>
	/// Стандартное количество точек на дюйм
	/// </summary>
	internal static readonly Dpi Default = new Dpi(96.0);
	
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="Dpi"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Dpi(double value)
	{
		Value = value;
	}
}
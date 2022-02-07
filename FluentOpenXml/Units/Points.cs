namespace FluentOpenXml.Units;

/// <summary>
/// Представляет точки
/// </summary>
internal readonly struct Points
{
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="Points"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Points(double value)
	{
		Value = value;
	}
	
	/// <summary>
	/// Преобразует <see cref="Points"/> в <see cref="Emu"/>
	/// </summary>
	internal Emu ToEmu() => new Emu(Value * 12700.0);
	
	/// <summary>
	/// Преобразует <see cref="Points"/> в <see cref="Dxa"/>
	/// </summary>
	internal Dxa ToDxa() => new Dxa(Value * 20.0);
	
	/// <summary>
	/// Преобразует <see cref="Points"/> в <see cref="Centimeters"/>
	/// </summary>
	internal Centimeters ToCentimeters() => ToEmu().ToCentimeters();
	
	/// <summary>
	/// Преобразует <see cref="Points"/> в <see cref="Millimeters"/>
	/// </summary>
	internal Millimeters ToMillimeters() => ToCentimeters().ToMillimeters();
	
	/// <summary>
	/// Преобразует <see cref="Points"/> в <see cref="HalfPoints"/>
	/// </summary>
	internal HalfPoints ToHalfPoints() => new HalfPoints(Value * 2.0);
	
	/// <summary>
	/// Преобразует <see cref="Points"/> в <see cref="Pixels"/>, используя указанное количество точек на дюйм
	/// </summary>
	/// <param name="dpi">Количество точек на дюйм</param>
	internal Pixels ToPixels(Dpi dpi) => new Pixels(Value / 72.0, dpi);

	/// <summary>
	/// Преобразует <see cref="Points"/> в <see cref="Pixels"/>, используя стандартное количество точек на дюйм
	/// </summary>
	internal Pixels ToPixels() => ToPixels(Dpi.Default);
}
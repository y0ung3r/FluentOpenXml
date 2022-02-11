namespace FluentOpenXml.Units;

/// <summary>
/// Представляет английские единицы измерения
/// https://en.wikipedia.org/wiki/Office_Open_XML_file_formats
/// https://en.wikipedia.org/wiki/English_units
/// </summary>
internal readonly struct Emu
{
	/// <summary>
	/// Значение
	/// </summary>
	internal long Value { get; }

	/// <summary>
	/// Инициализирует <see cref="Emu"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Emu(double value)
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
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Inches"/>
	/// </summary>
	internal Inches ToInches() => new Inches(Value / 914400.0);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Centimeters"/>
	/// </summary>
	internal Centimeters ToCentimeters() => new Centimeters(Value / 360000.0);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Dxa"/>
	/// </summary>
	internal Dxa ToDxa() => new Dxa(Value / 635.0);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Millimeters"/>
	/// </summary>
	internal Millimeters ToMillimeters() => new Millimeters(Value / 36000.0);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Points"/>
	/// </summary>
	internal Points ToPoints() => new Points(Value / 12700.0);

	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Pixels"/>, используя указанное количество точек на дюйм
	/// </summary>
	/// <param name="dpi">Количество точек на дюйм</param>
	internal Pixels ToPixel(Dpi dpi) => new Pixels(Value / 914400.0, dpi);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Pixels"/>, используя стандартное количество точек на дюйм
	/// </summary>
	internal Pixels ToPixel() => ToPixel(Dpi.Default);
}
namespace FluentOpenXml.Units;

/// <summary>
/// Представляет дюймы
/// </summary>
internal readonly struct Inches
{
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="Inches"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Inches(double value)
	{
		if (value < 0)
		{
			throw new ArgumentException($"Значение \"{nameof(value)}\" не может быть меньше нуля");
		}
		
		Value = value;
	}
	
	/// <summary>
	/// Преобразует <see cref="Inches"/> в <see cref="Emu"/>
	/// </summary>
	internal Emu ToEmu() => new Emu(Value * 914400.0);
	
	/// <summary>
	/// Преобразует <see cref="Inches"/> в <see cref="Dxa"/>
	/// </summary>
	internal Dxa ToDxa() => new Dxa(Value * 72.0);

	/// <summary>
	/// Преобразует <see cref="Inches"/> в <see cref="Pixels"/>, используя указанное количество точек на дюйм
	/// </summary>
	/// <param name="dpi">Количество точек на дюйм</param>
	internal Pixels ToPixels(Dpi dpi) => ToEmu().ToPixel(dpi);

	/// <summary>
	/// Преобразует <see cref="Inches"/> в <see cref="Pixels"/>, используя стандартное количество точек на дюйм
	/// </summary>
	internal Pixels ToPixels() => ToPixels(Dpi.Default);
}
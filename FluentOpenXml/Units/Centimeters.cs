namespace FluentOpenXml.Units;

/// <summary>
/// Представляет сантиметры
/// </summary>
internal readonly struct Centimeters
{
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="Centimeters"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Centimeters(double value)
	{
		if (value < 0)
		{
			throw new ArgumentException($"Значение \"{nameof(value)}\" не может быть меньше нуля");
		}
		
		Value = value;
	}
	
	/// <summary>
	/// Преобразует <see cref="Centimeters"/> в <see cref="Emu"/>
	/// </summary>
	internal Emu ToEmu() => new Emu(Value * 360000.0);
	
	/// <summary>
	/// Преобразует <see cref="Centimeters"/> в <see cref="Millimeters"/>
	/// </summary>
	/// <returns></returns>
	internal Millimeters ToMillimeters() => new Millimeters(Value * 10.0);
	
	/// <summary>
	/// Преобразует <see cref="Centimeters"/> в <see cref="Points"/>
	/// </summary>
	internal Points ToPoints() => ToEmu().ToPoints();
}
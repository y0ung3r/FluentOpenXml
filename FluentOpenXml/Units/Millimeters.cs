namespace FluentOpenXml.Units;

/// <summary>
/// Представляет миллиметры
/// </summary>
internal readonly struct Millimeters
{
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="Millimeters"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Millimeters(double value)
	{
		if (value < 0)
		{
			throw new ArgumentException($"Значение \"{nameof(value)}\" не может быть меньше нуля");
		}
		
		Value = value;
	}
	
	/// <summary>
	/// Преобразует <see cref="Millimeters"/> в <see cref="Emu"/>
	/// </summary>
	internal Emu ToEmu() => new Emu(Value * 36000.0);
	
	/// <summary>
	/// Преобразует <see cref="Millimeters"/> в <see cref="Centimeters"/>
	/// </summary>
	internal Centimeters ToCentimeters() => new Centimeters(Value / 10.0);
	
	/// <summary>
	/// Преобразует <see cref="Millimeters"/> в <see cref="Points"/>
	/// </summary>
	internal Points ToPoints() => ToCentimeters().ToPoints();
}
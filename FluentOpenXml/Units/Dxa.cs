namespace FluentOpenXml.Units;

/// <summary>
/// Представляет двадцатые доли точки
/// Используется в отступах, интервалах и т. д.
/// </summary>
internal readonly struct Dxa
{
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="Dxa"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Dxa(double value)
	{
		if (value < 0)
		{
			throw new ArgumentException($"Значение \"{nameof(value)}\" не может быть меньше нуля");
		}
		
		Value = value;
	}
	
	/// <summary>
	/// Преобразует <see cref="Dxa"/> в <see cref="Emu"/>
	/// </summary>
	internal Emu ToEmu() => new Emu(Value * 635.0);
	
	/// <summary>
	/// Преобразует <see cref="Dxa"/> в <see cref="Points"/>
	/// </summary>
	internal Points ToPoints() => new Points(Value / 20.0);
	
	/// <summary>
	/// Преобразует <see cref="Dxa"/> в <see cref="Inches"/>
	/// </summary>
	internal Inches ToInches() => new Inches(Value / 72.0);
}
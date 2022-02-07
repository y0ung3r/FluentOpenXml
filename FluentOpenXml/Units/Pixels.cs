namespace FluentOpenXml.Units;

/// <summary>
/// Представляет пиксели
/// </summary>
internal readonly struct Pixels
{
	/// <summary>
	/// Значение с учетом количества точек на дюйм
	/// </summary>
	internal double Value { get; }
	
	/// <summary>
	/// Количество точек на дюйм, использованные для расчета пикселей
	/// </summary>
	internal Dpi Dpi { get; }

	/// <summary>
	/// Инициализирует <see cref="Pixels"/>
	/// </summary>
	/// <param name="value">Значение</param>
	/// <param name="dpi">Количество точек на дюйм, использованные для расчета пикселей</param>
	internal Pixels(double value, Dpi dpi)
	{
		Dpi = dpi;
		Value = value * Dpi.Value;
	}
	
	/// <summary>
	/// Преобразует <see cref="Pixels"/> в <see cref="Emu"/>
	/// </summary>
	internal Emu ToEmu() => new Emu(Value * 914400.0 / Dpi.Value);
	
	/// <summary>
	/// Преобразует <see cref="Pixels"/> в <see cref="Points"/>
	/// </summary>
	/// <returns></returns>
	internal Points ToPoints() => new Points(Value * 72.0 / Dpi.Value);
	
	/// <summary>
	/// Преобразует <see cref="Pixels"/> в <see cref="Inches"/>
	/// </summary>
	internal Inches ToInches() => ToEmu().ToInches();
}
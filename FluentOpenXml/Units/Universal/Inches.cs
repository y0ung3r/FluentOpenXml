using FluentOpenXml.Units.OpenXml;

namespace FluentOpenXml.Units.Universal;

/// <summary>
/// Представляет дюймы
/// </summary>
public sealed class Inches : UniversalUnits
{
	/// <summary>
	/// Инициализирует <see cref="Inches"/>
	/// </summary>
	/// <param name="value">Значение</param>
	public Inches(double value) 
		: base(value)
	{ }
	
	/// <summary>
	/// Преобразует <see cref="Inches"/> в <see cref="Emu"/>
	/// </summary>
	internal override Emu ToEmu() => new Emu(Value * 914400.0);
}
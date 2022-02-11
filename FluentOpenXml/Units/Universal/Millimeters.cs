using FluentOpenXml.Units.OpenXml;

namespace FluentOpenXml.Units.Universal;

/// <summary>
/// Представляет миллиметры
/// </summary>
public sealed class Millimeters : UniversalUnits
{
	/// <summary>
	/// Инициализирует <see cref="Millimeters"/>
	/// </summary>
	/// <param name="value">Значение</param>
	public Millimeters(double value) 
		: base(value)
	{ }
	
	/// <summary>
	/// Преобразует <see cref="Millimeters"/> в <see cref="Emu"/>
	/// </summary>
	internal override Emu ToEmu() => new Emu(Value * 36000.0);
}
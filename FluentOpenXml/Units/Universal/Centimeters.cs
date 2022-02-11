using FluentOpenXml.Units.OpenXml;

namespace FluentOpenXml.Units.Universal;

/// <summary>
/// Представляет сантиметры
/// </summary>
public sealed class Centimeters : UniversalUnits
{
	/// <summary>
	/// Инициализирует <see cref="Centimeters"/>
	/// </summary>
	/// <param name="value">Значение</param>
	public Centimeters(double value) 
		: base(value)
	{ }
	
	/// <summary>
	/// Преобразует <see cref="Centimeters"/> в <see cref="Emu"/>
	/// </summary>
	internal override Emu ToEmu() => new Emu(Value * 360000.0);
}
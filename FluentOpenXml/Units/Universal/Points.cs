using FluentOpenXml.Units.OpenXml;

namespace FluentOpenXml.Units.Universal;

/// <summary>
/// Представляет пункты
/// </summary>
public sealed class Points : UniversalUnits
{
	/// <summary>
	/// Инициализирует <see cref="Points"/>
	/// </summary>
	/// <param name="value">Значение</param>
	public Points(double value) 
		: base(value)
	{ }

	/// <summary>
	/// Преобразует <see cref="Points"/> в <see cref="HalfPoints"/>
	/// </summary>
	internal HalfPoints ToHalfPoints() => new HalfPoints(Value * 2.0);
	
	/// <summary>
	/// Преобразует <see cref="Points"/> в <see cref="Emu"/>
	/// </summary>
	internal override Emu ToEmu() => new Emu(Value * 12700.0);
}
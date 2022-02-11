using FluentOpenXml.Units.OpenXml;

namespace FluentOpenXml.Units.Universal;

/// <summary>
/// Представляет половины от пункта
/// </summary>
internal sealed class HalfPoints : UniversalUnits
{
	/// <summary>
	/// Инициализирует <see cref="HalfPoints"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal HalfPoints(double value)
		: base(value)
	{ }

	/// <summary>
	/// Преобразует <see cref="HalfPoints"/> в <see cref="Points"/>
	/// </summary>
	internal Points ToPoints() => new Points(Value / 2.0);
	
	/// <summary>
	/// Преобразует <see cref="HalfPoints"/> в <see cref="Emu"/>
	/// </summary>
	/// <returns></returns>
	internal override Emu ToEmu() => ToPoints().ToEmu();
}
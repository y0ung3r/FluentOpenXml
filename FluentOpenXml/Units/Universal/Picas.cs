using FluentOpenXml.Units.OpenXml;

namespace FluentOpenXml.Units.Universal;

/// <summary>
/// Представляет пики
/// </summary>
public sealed class Picas : UniversalUnits
{
	/// <summary>
	/// Инициализирует <see cref="Picas"/>
	/// </summary>
	/// <param name="value">Значение</param>
	public Picas(double value)
		: base(value)
	{ }

	/// <summary>
	/// Преобразовывает <see cref="Picas"/> в <see cref="Inches"/>
	/// </summary>
	/// <returns></returns>
	internal Inches ToInches() => new Inches(Value / 6.0);

	/// <summary>
	/// Преобразовывает <see cref="Picas"/> в <see cref="Emu"/>
	/// </summary>
	internal override Emu ToEmu() => ToInches().ToEmu();
}
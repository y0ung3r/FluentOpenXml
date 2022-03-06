namespace FluentOpenXml.Units.Universal;

/// <summary>
/// Представляет базовый класс для универсальных единиц измерения
/// </summary>
public abstract class UniversalUnits : FloatingPointUnits
{
	/// <summary>
	/// Инициализирует <see cref="UniversalUnits"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal UniversalUnits(double value)
		: base(value)
	{ }

	/// <summary>
	/// Преобразовывает <see cref="UniversalUnits"/> в <see cref="Emu"/>
	/// </summary>
	internal abstract Emu ToEmu();
}
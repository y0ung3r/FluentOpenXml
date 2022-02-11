namespace FluentOpenXml.Units.OpenXml;

/// <summary>
/// Представляет двадцатые доли пункта
/// </summary>
internal sealed class Twips : OpenXmlUnits
{
	/// <summary>
	/// Инициализирует <see cref="Twips"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Twips(double value)
		: base(value)
	{ }
	
	/// <summary>
	/// Преобразует <see cref="Twips"/> в <see cref="Emu"/>
	/// </summary>
	internal Emu ToEmu() => new Emu(Value * 635.0);
}
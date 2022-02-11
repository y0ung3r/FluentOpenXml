using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Units.OpenXml;

/// <summary>
/// Представляет английские единицы измерения
/// https://en.wikipedia.org/wiki/Office_Open_XML_file_formats
/// https://en.wikipedia.org/wiki/English_units
/// </summary>
internal sealed class Emu : OpenXmlUnits
{
	/// <summary>
	/// Инициализирует <see cref="Emu"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal Emu(double value)
		: base(value)
	{ }
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Centimeters"/>
	/// </summary>
	internal Centimeters ToCentimeters() => new Centimeters(Value / 360000.0);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Points"/>
	/// </summary>
	internal Points ToPoints() => new Points(Value / 12700.0);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Inches"/>
	/// </summary>
	internal Inches ToInches() => new Inches(Value / 914400.0);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Millimeters"/>
	/// </summary>
	internal Millimeters ToMillimeters() => new Millimeters(Value / 36000.0);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Twips"/>
	/// </summary>
	internal Twips ToTwips() => new Twips(Value / 635.0);
}
using System.Reflection;
using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Units;

/// <summary>
/// Представляет английские единицы измерения
/// https://en.wikipedia.org/wiki/Office_Open_XML_file_formats
/// </summary>
internal sealed class Emu : FloatingPointUnits
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
	/// Преобразует <see cref="Emu"/> в <see cref="Picas"/>
	/// </summary>
	internal Picas ToPicas() => new Picas(Value / 152400.0);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Millimeters"/>
	/// </summary>
	internal Millimeters ToMillimeters() => new Millimeters(Value / 36000.0);
	
	/// <summary>
	/// Преобразует <see cref="Emu"/> в <see cref="Twips"/>
	/// </summary>
	internal Twips ToTwips() => new Twips(Value / 635.0);

	/// <summary>
	/// Преобразует <see cref="Emu"/> в указанную универсальную единицу измерения <see cref="TUnits"/>
	/// </summary>
	/// <typeparam name="TUnits">Единица измерения</typeparam>
	internal TUnits To<TUnits>()
		where TUnits : UniversalUnits
	{
		var methods = GetType().GetMethods(BindingFlags.Instance | BindingFlags.NonPublic);
		var conversionMethod = methods.FirstOrDefault
		(
			method => method.ReturnType == typeof(TUnits)
		);

		if (conversionMethod is null)
		{
			throw new ArgumentException($"Метод преобразования для \"{typeof(TUnits).Name}\" не найден");
		}

		return (TUnits)conversionMethod.Invoke
		(
			this, 
			Array.Empty<object>()
		);
	}
}
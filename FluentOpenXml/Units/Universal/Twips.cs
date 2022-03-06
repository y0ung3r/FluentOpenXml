using DocumentFormat.OpenXml;

namespace FluentOpenXml.Units.Universal;

/// <summary>
/// Представляет двадцатые доли пункта
/// </summary>
internal sealed class Twips : UniversalUnits
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
	internal override Emu ToEmu() => new Emu(Value * 635.0);
	
	/// <summary>
	/// Преобразует <see cref="Twips"/> в <see cref="UInt16Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="Twips"/></param>
	public static implicit operator UInt16Value(Twips units) => Convert.ToUInt16(units.Value);
	
	/// <summary>
	/// Преобразует <see cref="Twips"/> в <see cref="UInt32Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="Twips"/></param>
	public static implicit operator UInt32Value(Twips units) => Convert.ToUInt32(units.Value);
	
	/// <summary>
	/// Преобразует <see cref="Twips"/> в <see cref="UInt64Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="Twips"/></param>
	public static implicit operator UInt64Value(Twips units) => Convert.ToUInt64(units.Value);
	
	/// <summary>
	/// Преобразует <see cref="Twips"/> в <see cref="Int16Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="Twips"/></param>
	public static implicit operator Int16Value(Twips units) => Convert.ToInt16(units.Value);
	
	/// <summary>
	/// Преобразует <see cref="Twips"/> в <see cref="Int32Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="Twips"/></param>
	public static implicit operator Int32Value(Twips units) => Convert.ToInt32(units.Value);
	
	/// <summary>
	/// Преобразует <see cref="Twips"/> в <see cref="Int64Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="Twips"/></param>
	public static implicit operator Int64Value(Twips units) => Convert.ToInt64(units.Value);
}
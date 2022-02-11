using DocumentFormat.OpenXml;

namespace FluentOpenXml.Units.OpenXml;

/// <summary>
/// Представляет базовый класс для единиц измерения, используемых в Open XML
/// </summary>
internal abstract class OpenXmlUnits
{
	/// <summary>
	/// Значение
	/// </summary>
	internal long Value { get; }

	/// <summary>
	/// Инициализирует <see cref="OpenXmlUnits"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal OpenXmlUnits(double value)
	{
		if (value < 0)
		{
			throw new ArgumentException($"Значение \"{nameof(value)}\" не может быть меньше нуля");
		}
		
		Value = Convert.ToInt64
		(
			Math.Truncate(value)
		);
	}

	/// <summary>
	/// Преобразует <see cref="OpenXmlUnits"/> в <see cref="UInt16Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="OpenXmlUnits"/></param>
	public static implicit operator UInt16Value(OpenXmlUnits units) => Convert.ToUInt16(units.Value);
	
	/// <summary>
	/// Преобразует <see cref="OpenXmlUnits"/> в <see cref="UInt32Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="OpenXmlUnits"/></param>
	public static implicit operator UInt32Value(OpenXmlUnits units) => Convert.ToUInt32(units.Value);
	
	/// <summary>
	/// Преобразует <see cref="OpenXmlUnits"/> в <see cref="UInt64Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="OpenXmlUnits"/></param>
	public static implicit operator UInt64Value(OpenXmlUnits units) => Convert.ToUInt64(units.Value);
	
	/// <summary>
	/// Преобразует <see cref="OpenXmlUnits"/> в <see cref="Int16Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="OpenXmlUnits"/></param>
	public static implicit operator Int16Value(OpenXmlUnits units) => Convert.ToInt16(units.Value);
	
	/// <summary>
	/// Преобразует <see cref="OpenXmlUnits"/> в <see cref="Int32Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="OpenXmlUnits"/></param>
	public static implicit operator Int32Value(OpenXmlUnits units) => Convert.ToInt32(units.Value);
	
	/// <summary>
	/// Преобразует <see cref="OpenXmlUnits"/> в <see cref="Int64Value"/>
	/// </summary>
	/// <param name="units">Значение, представленное в виде <see cref="OpenXmlUnits"/></param>
	public static implicit operator Int64Value(OpenXmlUnits units) => Convert.ToInt64(units.Value);
}
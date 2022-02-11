using FluentOpenXml.Units.OpenXml;

namespace FluentOpenXml.Units.Universal;

/// <summary>
/// Представляет пункты
/// </summary>
public sealed class Points : UniversalUnits
{
	#region Стандартные значения

	/// <summary>
	/// Возвращает 8 пунктов
	/// </summary>
	public static Points P8 => new Points(8.0);
	
	/// <summary>
	/// Возвращает 9 пунктов
	/// </summary>
	public static Points P9 => new Points(9.0);
	
	/// <summary>
	/// Возвращает 10 пунктов
	/// </summary>
	public static Points P10 => new Points(10.0);
	
	/// <summary>
	/// Возвращает 11 пунктов
	/// </summary>
	public static Points P11 => new Points(11.0);
	
	/// <summary>
	/// Возвращает 12 пунктов
	/// </summary>
	public static Points P12 => new Points(12.0);
	
	/// <summary>
	/// Возвращает 14 пунктов
	/// </summary>
	public static Points P14 => new Points(14.0);
	
	/// <summary>
	/// Возвращает 16 пунктов
	/// </summary>
	public static Points P16 => new Points(16.0);
	
	/// <summary>
	/// Возвращает 18 пунктов
	/// </summary>
	public static Points P18 => new Points(18.0);
	
	/// <summary>
	/// Возвращает 20 пунктов
	/// </summary>
	public static Points P20 => new Points(20.0);
	
	/// <summary>
	/// Возвращает 22 пункта
	/// </summary>
	public static Points P22 => new Points(22.0);
	
	/// <summary>
	/// Возвращает 24 пункта
	/// </summary>
	public static Points P24 => new Points(24.0);
	
	/// <summary>
	/// Возвращает 26 пунктов
	/// </summary>
	public static Points P26 => new Points(26.0);
	
	/// <summary>
	/// Возвращает 28 пунктов
	/// </summary>
	public static Points P28 => new Points(28.0);
	
	/// <summary>
	/// Возвращает 36 пунктов
	/// </summary>
	public static Points P36 => new Points(36.0);
	
	/// <summary>
	/// Возвращает 48 пунктов
	/// </summary>
	public static Points P48 => new Points(48.0);
	
	/// <summary>
	/// Возвращает 72 пункта
	/// </summary>
	public static Points P72 => new Points(72.0);
	
	#endregion
	
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
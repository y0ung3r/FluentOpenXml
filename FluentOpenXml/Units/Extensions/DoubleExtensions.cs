using FluentOpenXml.Units.Universal;

namespace FluentOpenXml.Units.Extensions;

/// <summary>
/// Методы-расширения для <see cref="Double"/>
/// </summary>
internal static class DoubleExtensions
{
	/// <summary>
	/// Создает <see cref="TUnits"/> из <see cref="Double"/>
	/// </summary>
	/// <param name="value">Значение</param>
	/// <typeparam name="TUnits">Тип единицы измерения</typeparam>
	/// <exception cref="ArgumentException">Невозможно преобразовать указанное значение в <see cref="TUnits"/></exception>
	internal static TUnits As<TUnits>(this double value)
		where TUnits : FloatingPointUnits
	{
		ArgumentNullException.ThrowIfNull(value);

		try
		{
			return (TUnits)Activator.CreateInstance
			(
				typeof(TUnits), 
				value
			);
		}
		
		catch (MissingMethodException exception)
		{
			throw new ArgumentException($"Невозможно преобразовать \"{nameof(value)}\" в \"{typeof(TUnits).Name}\"", exception);
		}
	}

	/// <summary>
	/// Создает <see cref="Double"/> и выполняет преобразование в <see cref="Twips"/>
	/// </summary>
	/// <param name="value">Значение</param>
	/// <typeparam name="TUnits">Единица измерения</typeparam>
	internal static Twips ToTwipsAs<TUnits>(this double value)
		where TUnits : UniversalUnits => value.As<TUnits>().ToEmu().ToTwips();
}
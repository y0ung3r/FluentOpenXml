namespace FluentOpenXml.Units;

/// <summary>
/// Представляет базовый класс для единиц измерения с плавающей точкой
/// </summary>
public abstract class FloatingPointUnits
{
	/// <summary>
	/// Создает экземлпяр единицы измерения с указанным значением и типом
	/// </summary>
	/// <param name="value">Значение</param>
	/// <typeparam name="TUnit">Тип единицы измерения, ограниченный по <see cref="FloatingPointUnits"/></typeparam>
	public static TUnit From<TUnit>(double value) 
		where TUnit : FloatingPointUnits => (TUnit)Activator.CreateInstance(typeof(TUnit), value);
	
	/// <summary>
	/// Значение
	/// </summary>
	internal double Value { get; }

	/// <summary>
	/// Инициализирует <see cref="FloatingPointUnits"/>
	/// </summary>
	/// <param name="value">Значение</param>
	internal FloatingPointUnits(double value)
	{
		if (value < 0)
		{
			throw new ArgumentException($"Значение \"{nameof(value)}\" не может быть меньше нуля");
		}
		
		Value = value;
	}
}
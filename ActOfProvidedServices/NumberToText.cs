//
// Классы для преобразования чисел и денежных сумм в пропись.
// Источник - http://morpher.ru/SummaPropisyu.aspx
//
// Данная реализация примечательна тем, что в ней почти нет 
// операций сложения строк, всё делается через StringBuilder.
// Работает в 1,5 раза быстрее своих аналогов на C#.
//
// Аналоги:
// http://www.gotdotnet.ru/Downloads/Examples/772.aspx (C#, Уфимцев)
// http://rsdn.ru/article/files/dotnet/RusNumber.xml (C#)
// http://sources.ru/builder/faq/031.html (C++)
// http://delphiworld.narod.ru/base/sum_written_out9.html (Pascal)
//
// Метод Пропись перегружен для сумм типа decimal и double.
// Все неотрицательные суммы типа decimal с не более чем двумя цифрами
// после запятой представимы в виде прописи.
// Для сумм типа double максимальное представимое число - Число.MaxDouble.
//

using System;
using System.Text;

namespace Slepov.Russian.СуммаПрописью {
	/// <summary>
	/// Класс для записи денежных сумм прописью: "тысяча рублей 00 копеек".
	/// </summary>
	/// <example>
	/// Сумма.Пропись (100, Валюта.Рубли); // "сто рублей 00 копеек"
	/// Валюта.Рубли.Пропись (123.45); // "сто двадцать три рубля 45 копеек"
	/// </example>
	public static class Сумма {
		/// <summary>
		/// Записывает пропись суммы в заданной валюте в <paramref name="result"/> строчными буквами.
		/// </summary>
		public static StringBuilder Пропись(decimal сумма, Валюта валюта, StringBuilder result) {
			decimal целая = Math.Floor(сумма);
			uint дробная = (uint)((сумма - целая) * 100);

			Число.Пропись(целая, валюта.ОсновнаяЕдиница, result);
			return ДобавитьКопейки(дробная, валюта, result);
		}

		/// <summary>
		/// Записывает пропись суммы в заданной валюте в <paramref name="result"/> строчными буквами.
		/// </summary>
		public static StringBuilder Пропись(double сумма, Валюта валюта, StringBuilder result) {
			double целая = Math.Floor(сумма);

			// Вынесение 100 за скобки позволяет избежать ошибки округления
			// например, когда сумма = 1234.51.
			uint дробная = (uint)(сумма * 100) - (uint)(целая * 100);

			Число.Пропись(целая, валюта.ОсновнаяЕдиница, result);
			return ДобавитьКопейки(дробная, валюта, result);
		}

		private static StringBuilder ДобавитьКопейки(uint дробная, Валюта валюта, StringBuilder result) {
			result.Append(' ');

			// Эта строчка выполняется быстрее, чем следующая за ней закомментированная.
			result.Append(дробная.ToString("00"));
			//result.AppendFormat ("{0:00}", дробная);

			result.Append(' ');
			result.Append(Число.Согласовать(валюта.ДробнаяЕдиница, дробная));

			return result;
		}

		/// <summary>
		/// Проверяет, подходит ли число для передачи методу 
		/// <see cref="Сумма.Пропись (decimal, Валюта)"/>.
		/// </summary>
		/// <remarks>
		/// Сумма должна быть неотрицательной и должна содержать 
		/// не более двух цифр после запятой.
		/// </remarks>
		/// <returns>
		/// Описание нарушенного ограничения или null.
		/// </returns>
		public static string ПроверитьСумму(decimal сумма) {
			if (сумма < 0) return "Сумма должна быть неотрицательной.";

			decimal целая = Math.Floor(сумма);
			decimal дробная = (сумма - целая) * 100;

			if (Math.Floor(дробная) != дробная) {
				return "Сумма должна содержать не более двух цифр после запятой.";
			}

			return null;
		}

		/// <summary>
		/// Возвращает пропись заданной суммы строчными буквами.
		/// </summary>
		public static string Пропись(decimal n, Валюта валюта) {
			return Число.ApplyCaps(Пропись(n, валюта, new StringBuilder()), Заглавные.Нет);
		}

		/// <summary>
		/// Возвращает пропись заданной суммы.
		/// </summary>
		public static string Пропись(decimal n, Валюта валюта, Заглавные заглавные) {
			return Число.ApplyCaps(Пропись(n, валюта, new StringBuilder()), заглавные);
		}

		/// <summary>
		/// Возвращает пропись заданной суммы строчными буквами.
		/// </summary>
		public static string Пропись(double n, Валюта валюта) {
			return Число.ApplyCaps(Пропись(n, валюта, new StringBuilder()), Заглавные.Нет);
		}

		/// <summary>
		/// Возвращает пропись заданной суммы.
		/// </summary>
		public static string Пропись(double n, Валюта валюта, Заглавные заглавные) {
			return Число.ApplyCaps(Пропись(n, валюта, new StringBuilder()), заглавные);
		}
	}

	/// <summary>
	/// Класс для преобразования чисел в пропись на русском языке.
	/// </summary>
	/// <example>
	/// Число.Пропись (1, РодЧисло.Мужской); // "один"
	/// Число.Пропись (2, РодЧисло.Женский); // "две"
	/// Число.Пропись (21, РодЧисло.Средний); // "двадцать одно"
	/// </example>
	/// <example>
	/// Число.Пропись (5, new ЕдиницаИзмерения (
	///  РодЧисло.Мужской, "метр", "метра", "метров"), sb); // "пять метров"
	/// </example>
	public static class Число {
		/// <summary>
		/// Получить пропись числа с согласованной единицей измерения.
		/// </summary>
		/// <param name="число"> Число должно быть целым, неотрицательным. </param>
		/// <param name="еи"></param>
		/// <param name="result"> Сюда записывается результат. </param>
		/// <returns> <paramref name="result"/> </returns>
		/// <exception cref="ArgumentException">
		/// Если число меньше нуля или не целое. 
		/// </exception>
		public static StringBuilder Пропись(decimal число, IЕдиницаИзмерения еи, StringBuilder result) {
			string error = ПроверитьЧисло(число);
			if (error != null) throw new ArgumentException(error, "число");

			// Целочисленные версии работают в разы быстрее, чем decimal.
			if (число <= uint.MaxValue) {
				Пропись((uint)число, еи, result);
			} else if (число <= ulong.MaxValue) {
				Пропись((ulong)число, еи, result);
			} else {
				MyStringBuilder mySb = new MyStringBuilder(result);

				decimal div1000 = Math.Floor(число / 1000);
				ПрописьСтаршихКлассов(div1000, 0, mySb);
				ПрописьКласса((uint)(число - div1000 * 1000), еи, mySb);
			}

			return result;
		}

		/// <summary>
		/// Получить пропись числа с согласованной единицей измерения.
		/// </summary>
		/// <param name="число"> 
		/// Число должно быть целым, неотрицательным, не большим <see cref="MaxDouble"/>. 
		/// </param>
		/// <param name="еи"></param>
		/// <param name="result"> Сюда записывается результат. </param>
		/// <exception cref="ArgumentException">
		/// Если число меньше нуля, не целое или больше <see cref="MaxDouble"/>. 
		/// </exception>
		/// <returns> <paramref name="result"/> </returns>
		/// <remarks>
		/// float по умолчанию преобразуется к double, поэтому нет перегрузки для float.
		/// В результате ошибок округления возможно расхождение цифр прописи и
		/// строки, выдаваемой double.ToString ("R"), начиная с 17 значащей цифры.
		/// </remarks>
		public static StringBuilder Пропись(double число, IЕдиницаИзмерения еи, StringBuilder result) {
			string error = ПроверитьЧисло(число);
			if (error != null) throw new ArgumentException(error, "число");

			if (число <= uint.MaxValue) {
				Пропись((uint)число, еи, result);
			} else if (число <= ulong.MaxValue) {
				// Пропись с ulong выполняется в среднем в 2 раза быстрее.
				Пропись((ulong)число, еи, result);
			} else {
				MyStringBuilder mySb = new MyStringBuilder(result);

				double div1000 = Math.Floor(число / 1000);
				ПрописьСтаршихКлассов(div1000, 0, mySb);
				ПрописьКласса((uint)(число - div1000 * 1000), еи, mySb);
			}

			return result;
		}

		/// <summary>
		/// Получить пропись числа с согласованной единицей измерения.
		/// </summary>
		/// <returns> <paramref name="result"/> </returns>
		public static StringBuilder Пропись(ulong число, IЕдиницаИзмерения еи, StringBuilder result) {
			if (число <= uint.MaxValue) {
				Пропись((uint)число, еи, result);
			} else {
				MyStringBuilder mySb = new MyStringBuilder(result);

				ulong div1000 = число / 1000;
				ПрописьСтаршихКлассов(div1000, 0, mySb);
				ПрописьКласса((uint)(число - div1000 * 1000), еи, mySb);
			}

			return result;
		}

		/// <summary>
		/// Получить пропись числа с согласованной единицей измерения.
		/// </summary>
		/// <returns> <paramref name="result"/> </returns>
		public static StringBuilder Пропись(uint число, IЕдиницаИзмерения еи, StringBuilder result) {
			MyStringBuilder mySb = new MyStringBuilder(result);

			if (число == 0) {
				mySb.Append("ноль");
				mySb.Append(еи.РодМнож);
			} else {
				uint div1000 = число / 1000;
				ПрописьСтаршихКлассов(div1000, 0, mySb);
				ПрописьКласса(число - div1000 * 1000, еи, mySb);
			}

			return result;
		}

		/// <summary>
		/// Записывает в <paramref name="sb"/> пропись числа, начиная с самого 
		/// старшего класса до класса с номером <paramref name="номерКласса"/>.
		/// </summary>
		/// <param name="sb"></param>
		/// <param name="число"></param>
		/// <param name="номерКласса">0 = класс тысяч, 1 = миллионов и т.д.</param>
		/// <remarks>
		/// В методе применена рекурсия, чтобы обеспечить запись в StringBuilder 
		/// в нужном порядке - от старших классов к младшим.
		/// </remarks>
		static void ПрописьСтаршихКлассов(decimal число, int номерКласса, MyStringBuilder sb) {
			if (число == 0) return; // конец рекурсии

			// Записать в StringBuilder пропись старших классов.
			decimal div1000 = Math.Floor(число / 1000);
			ПрописьСтаршихКлассов(div1000, номерКласса + 1, sb);

			uint числоДо999 = (uint)(число - div1000 * 1000);
			if (числоДо999 == 0) return;

			ПрописьКласса(числоДо999, Классы[номерКласса], sb);
		}

		static void ПрописьСтаршихКлассов(double число, int номерКласса, MyStringBuilder sb) {
			if (число == 0) return; // конец рекурсии

			// Записать в StringBuilder пропись старших классов.
			double div1000 = Math.Floor(число / 1000);
			ПрописьСтаршихКлассов(div1000, номерКласса + 1, sb);

			uint числоДо999 = (uint)(число - div1000 * 1000);
			if (числоДо999 == 0) return;

			ПрописьКласса(числоДо999, Классы[номерКласса], sb);
		}

		static void ПрописьСтаршихКлассов(ulong число, int номерКласса, MyStringBuilder sb) {
			if (число == 0) return; // конец рекурсии

			// Записать в StringBuilder пропись старших классов.
			ulong div1000 = число / 1000;
			ПрописьСтаршихКлассов(div1000, номерКласса + 1, sb);

			uint числоДо999 = (uint)(число - div1000 * 1000);
			if (числоДо999 == 0) return;

			ПрописьКласса(числоДо999, Классы[номерКласса], sb);
		}

		static void ПрописьСтаршихКлассов(uint число, int номерКласса, MyStringBuilder sb) {
			if (число == 0) return; // конец рекурсии

			// Записать в StringBuilder пропись старших классов.
			uint div1000 = число / 1000;
			ПрописьСтаршихКлассов(div1000, номерКласса + 1, sb);

			uint числоДо999 = число - div1000 * 1000;
			if (числоДо999 == 0) return;

			ПрописьКласса(числоДо999, Классы[номерКласса], sb);
		}

		#region ПрописьКласса

		/// <summary>
		/// Формирует запись класса с названием, например,
		/// "125 тысяч", "15 рублей".
		/// Для 0 записывает только единицу измерения в род.мн.
		/// </summary>
		private static void ПрописьКласса(uint числоДо999, IЕдиницаИзмерения класс, MyStringBuilder sb) {
			uint числоЕдиниц = числоДо999 % 10;
			uint числоДесятков = (числоДо999 / 10) % 10;
			uint числоСотен = (числоДо999 / 100) % 10;

			sb.Append(Сотни[числоСотен]);

			if ((числоДо999 % 100) != 0) {
				Десятки[числоДесятков].Пропись(sb, числоЕдиниц, класс.РодЧисло);
			}

			// Добавить название класса в нужной форме.
			sb.Append(Согласовать(класс, числоДо999));
		}

		#endregion

		#region ПроверитьЧисло

		/// <summary>
		/// Проверяет, подходит ли число для передачи методу 
		/// <see cref="Пропись(decimal,IЕдиницаИзмерения,StringBuilder)"/>.
		/// </summary>
		/// <returns>
		/// Описание нарушенного ограничения или null.
		/// </returns>
		public static string ПроверитьЧисло(decimal число) {
			if (число < 0)
				return "Число должно быть больше или равно нулю.";

			if (число != decimal.Floor(число))
				return "Число не должно содержать дробной части.";

			return null;
		}

		/// <summary>
		/// Проверяет, подходит ли число для передачи методу 
		/// <see cref="Пропись(double,IЕдиницаИзмерения,StringBuilder)"/>.
		/// </summary>
		/// <returns>
		/// Описание нарушенного ограничения или null.
		/// </returns>
		public static string ПроверитьЧисло(double число) {
			if (число < 0)
				return "Число должно быть больше или равно нулю.";

			if (число != Math.Floor(число))
				return "Число не должно содержать дробной части.";

			if (число > MaxDouble) {
				return "Число должно быть не больше " + MaxDouble + ".";
			}

			return null;
		}

		#endregion

		#region Согласовать

		/// <summary>
		/// Согласовать название единицы измерения с числом.
		/// Например, согласование единицы (рубль, рубля, рублей) 
		/// с числом 23 даёт "рубля", а с числом 25 - "рублей".
		/// </summary>
		public static string Согласовать(IЕдиницаИзмерения единицаИзмерения, uint число) {
			uint числоЕдиниц = число % 10;
			uint числоДесятков = (число / 10) % 10;

			if (числоДесятков == 1) return единицаИзмерения.РодМнож;
			switch (числоЕдиниц) {
				case 1: return единицаИзмерения.ИменЕдин;
				case 2: case 3: case 4: return единицаИзмерения.РодЕдин;
				default: return единицаИзмерения.РодМнож;
			}
		}

		/// <summary>
		/// Согласовать название единицы измерения с числом.
		/// Например, согласование единицы (рубль, рубля, рублей) 
		/// с числом 23 даёт "рубля", а с числом 25 - "рублей".
		/// </summary>
		public static string Согласовать(IЕдиницаИзмерения единицаИзмерения, decimal число) {
			return Согласовать(единицаИзмерения, (uint)(число % 100));
		}

		#endregion

		#region Единицы

		static string ПрописьЦифры(uint цифра, РодЧисло род) {
			return Цифры[цифра].Пропись(род);
		}

		abstract class Цифра {
			public abstract string Пропись(РодЧисло род);
		}

		class ЦифраИзменяющаясяПоРодам : Цифра, IИзменяетсяПоРодам {
			public ЦифраИзменяющаясяПоРодам(
				string мужской,
				string женский,
				string средний,
				string множественное) {
				this.мужской = мужской;
				this.женский = женский;
				this.средний = средний;
				this.множественное = множественное;
			}

			public ЦифраИзменяющаясяПоРодам(
				string единственное,
				string множественное)

				: this(единственное, единственное, единственное, множественное) {
			}

			private readonly string мужской;
			private readonly string женский;
			private readonly string средний;
			private readonly string множественное;

			#region IИзменяетсяПоРодам Members

			public string Мужской { get { return this.мужской; } }
			public string Женский { get { return this.женский; } }
			public string Средний { get { return this.средний; } }
			public string Множественное { get { return this.множественное; } }

			#endregion

			public override string Пропись(РодЧисло род) {
				return род.ПолучитьФорму(this);
			}
		}

		class ЦифраНеизменяющаясяПоРодам : Цифра {
			public ЦифраНеизменяющаясяПоРодам(string пропись) {
				this.пропись = пропись;
			}

			private readonly string пропись;

			public override string Пропись(РодЧисло род) {
				return this.пропись;
			}
		}

		private static readonly Цифра[] Цифры = new Цифра[]
		{
			null,
			new ЦифраИзменяющаясяПоРодам ("один", "одна", "одно", "одни"),
			new ЦифраИзменяющаясяПоРодам ("два", "две", "два", "двое"),
			new ЦифраИзменяющаясяПоРодам ("три", "трое"),
			new ЦифраИзменяющаясяПоРодам ("четыре", "четверо"),
			new ЦифраНеизменяющаясяПоРодам ("пять"),
			new ЦифраНеизменяющаясяПоРодам ("шесть"),
			new ЦифраНеизменяющаясяПоРодам ("семь"),
			new ЦифраНеизменяющаясяПоРодам ("восемь"),
			new ЦифраНеизменяющаясяПоРодам ("девять"),
		};

		#endregion
		#region Десятки

		static readonly Десяток[] Десятки = new Десяток[]
		{
			new ПервыйДесяток (),
			new ВторойДесяток (),
			new ОбычныйДесяток ("двадцать"),
			new ОбычныйДесяток ("тридцать"),
			new ОбычныйДесяток ("сорок"),
			new ОбычныйДесяток ("пятьдесят"),
			new ОбычныйДесяток ("шестьдесят"),
			new ОбычныйДесяток ("семьдесят"),
			new ОбычныйДесяток ("восемьдесят"),
			new ОбычныйДесяток ("девяносто")
		};

		abstract class Десяток {
			public abstract void Пропись(MyStringBuilder sb, uint числоЕдиниц, РодЧисло род);
		}

		class ПервыйДесяток : Десяток {
			public override void Пропись(MyStringBuilder sb, uint числоЕдиниц, РодЧисло род) {
				sb.Append(ПрописьЦифры(числоЕдиниц, род));
			}
		}

		class ВторойДесяток : Десяток {
			static readonly string[] ПрописьНаДцать = new string[]
			{
				"десять",
				"одиннадцать",
				"двенадцать",
				"тринадцать",
				"четырнадцать",
				"пятнадцать",
				"шестнадцать",
				"семнадцать",
				"восемнадцать",
				"девятнадцать"
			};

			public override void Пропись(MyStringBuilder sb, uint числоЕдиниц, РодЧисло род) {
				sb.Append(ПрописьНаДцать[числоЕдиниц]);
			}
		}

		class ОбычныйДесяток : Десяток {
			public ОбычныйДесяток(string названиеДесятка) {
				this.названиеДесятка = названиеДесятка;
			}

			private readonly string названиеДесятка;

			public override void Пропись(MyStringBuilder sb, uint числоЕдиниц, РодЧисло род) {
				sb.Append(this.названиеДесятка);

				if (числоЕдиниц == 0) {
					// После "двадцать", "тридцать" и т.д. не пишут "ноль" (единиц)
				} else {
					sb.Append(ПрописьЦифры(числоЕдиниц, род));
				}
			}
		}

		#endregion
		#region Сотни

		static readonly string[] Сотни = new string[]
		{
			null,
			"сто",
			"двести",
			"триста",
			"четыреста",
			"пятьсот",
			"шестьсот",
			"семьсот",
			"восемьсот",
			"девятьсот"
		};

		#endregion
		#region Классы

		#region КлассТысяч

		class КлассТысяч : IЕдиницаИзмерения {
			public string ИменЕдин { get { return "тысяча"; } }
			public string РодЕдин { get { return "тысячи"; } }
			public string РодМнож { get { return "тысяч"; } }
			public РодЧисло РодЧисло { get { return РодЧисло.Женский; } }
		}

		#endregion
		#region Класс

		class Класс : IЕдиницаИзмерения {
			readonly string начальнаяФорма;

			public Класс(string начальнаяФорма) {
				this.начальнаяФорма = начальнаяФорма;
			}

			public string ИменЕдин { get { return this.начальнаяФорма; } }
			public string РодЕдин { get { return this.начальнаяФорма + "а"; } }
			public string РодМнож { get { return this.начальнаяФорма + "ов"; } }
			public РодЧисло РодЧисло { get { return РодЧисло.Мужской; } }
		}

		#endregion

		/// <summary>
		/// Класс - группа из 3 цифр.  Есть классы единиц, тысяч, миллионов и т.д.
		/// </summary>
		static readonly IЕдиницаИзмерения[] Классы = new IЕдиницаИзмерения[]
		{
			new КлассТысяч (),
			new Класс ("миллион"),
			new Класс ("миллиард"),
			new Класс ("триллион"),
			new Класс ("квадриллион"),
			new Класс ("квинтиллион"),
			new Класс ("секстиллион"),
			new Класс ("септиллион"),
			new Класс ("октиллион"),
 
            // Это количество классов покрывает весь диапазон типа decimal.
        };

		#endregion

		#region MaxDouble

		/// <summary>
		/// Максимальное число типа double, представимое в виде прописи.
		/// </summary>
		/// <remarks>
		/// Рассчитывается исходя из количества определённых классов.
		/// Если добавить ещё классы, оно будет автоматически увеличено.
		/// </remarks>
		public static double MaxDouble {
			get {
				if (maxDouble == 0) {
					maxDouble = CalcMaxDouble();
				}

				return maxDouble;
			}
		}

		private static double maxDouble = 0;

		static double CalcMaxDouble() {
			double max = Math.Pow(1000, Классы.Length + 1);

			double d = 1;

			while (max - d == max) {
				d *= 2;
			}

			return max - d;
		}

		#endregion

		#region Вспомогательные классы

		#region Форма

		#endregion
		#region MyStringBuilder

		/// <summary>
		/// Вспомогательный класс, аналогичный <see cref="StringBuilder"/>.
		/// Между вызовами <see cref="MyStringBuilder.Append"/> вставляет пробелы.
		/// </summary>
		class MyStringBuilder {
			public MyStringBuilder(StringBuilder sb) {
				this.sb = sb;
			}

			readonly StringBuilder sb;
			bool insertSpace = false;

			/// <summary>
			/// Добавляет слово <paramref name="s"/>,
			/// вставляя перед ним пробел, если нужно.
			/// </summary>
			public void Append(string s) {
				if (string.IsNullOrEmpty(s)) return;

				if (this.insertSpace) {
					this.sb.Append(' ');
				} else {
					this.insertSpace = true;
				}

				this.sb.Append(s);
			}

			public override string ToString() {
				return sb.ToString();
			}
		}

		#endregion

		#endregion

		#region Перегрузки метода Пропись, возвращающие string

		/// <summary>
		/// Возвращает пропись числа строчными буквами.
		/// </summary>
		public static string Пропись(decimal число, IЕдиницаИзмерения еи) {
			return Пропись(число, еи, Заглавные.Нет);
		}

		/// <summary>
		/// Возвращает пропись числа.
		/// </summary>
		public static string Пропись(decimal число, IЕдиницаИзмерения еи, Заглавные заглавные) {
			return ApplyCaps(Пропись(число, еи, new StringBuilder()), заглавные);
		}

		/// <summary>
		/// Возвращает пропись числа строчными буквами.
		/// </summary>
		public static string Пропись(double число, IЕдиницаИзмерения еи) {
			return Пропись(число, еи, Заглавные.Нет);
		}

		/// <summary>
		/// Возвращает пропись числа.
		/// </summary>
		public static string Пропись(double число, IЕдиницаИзмерения еи, Заглавные заглавные) {
			return ApplyCaps(Пропись(число, еи, new StringBuilder()), заглавные);
		}

		/// <summary>
		/// Возвращает пропись числа строчными буквами.
		/// </summary>
		public static string Пропись(ulong число, IЕдиницаИзмерения еи) {
			return Пропись(число, еи, Заглавные.Нет);
		}

		/// <summary>
		/// Возвращает пропись числа.
		/// </summary>
		public static string Пропись(ulong число, IЕдиницаИзмерения еи, Заглавные заглавные) {
			return ApplyCaps(Пропись(число, еи, new StringBuilder()), заглавные);
		}

		/// <summary>
		/// Возвращает пропись числа строчными буквами.
		/// </summary>
		public static string Пропись(uint число, IЕдиницаИзмерения еи) {
			return Пропись(число, еи, Заглавные.Нет);
		}

		/// <summary>
		/// Возвращает пропись числа.
		/// </summary>
		public static string Пропись(uint число, IЕдиницаИзмерения еи, Заглавные заглавные) {
			return ApplyCaps(Пропись(число, еи, new StringBuilder()), заглавные);
		}

		internal static string ApplyCaps(StringBuilder sb, Заглавные заглавные) {
			заглавные.Применить(sb);
			return sb.ToString();
		}

		#endregion
	}

	/// <summary>
	/// Стратегия расстановки заглавных букв.
	/// </summary>
	public abstract class Заглавные {
		/// <summary>
		/// Применить стратегию к <paramref name="sb"/>.
		/// </summary>
		public abstract void Применить(StringBuilder sb);

		class _ВСЕ : Заглавные {
			public override void Применить(StringBuilder sb) {
				for (int i = 0; i < sb.Length; ++i) {
					sb[i] = char.ToUpperInvariant(sb[i]);
				}
			}
		}

		class _Нет : Заглавные {
			public override void Применить(StringBuilder sb) {
			}
		}

		class _Первая : Заглавные {
			public override void Применить(StringBuilder sb) {
				sb[0] = char.ToUpperInvariant(sb[0]);
			}
		}

		public static readonly Заглавные ВСЕ = new _ВСЕ();
		public static readonly Заглавные Нет = new _Нет();
		public static readonly Заглавные Первая = new _Первая();
	}

	/// <summary>
	/// Описывает тип валюты как совокупность двух единиц измерения - основной и дробной.
	/// Содержит несколько предопределённых валют - рубли, доллары, евро.
	/// </summary>
	/// <remarks>
	/// Предполагается, что основная единица равна 100 дробным. 
	/// </remarks>
	public class Валюта {
		/// <summary> </summary>
		public Валюта(IЕдиницаИзмерения основная, IЕдиницаИзмерения дробная) {
			this.основная = основная;
			this.дробная = дробная;
		}

		readonly IЕдиницаИзмерения основная;
		readonly IЕдиницаИзмерения дробная;

		/// <summary>
		/// Основная единица измерения валюты - рубли, доллары, евро и т.д.
		/// </summary>
		public IЕдиницаИзмерения ОсновнаяЕдиница {
			get { return this.основная; }
		}

		/// <summary>
		/// Дробная единица измерения валюты - копейки, центы, евроценты и т.д.
		/// </summary>
		public IЕдиницаИзмерения ДробнаяЕдиница {
			get { return this.дробная; }
		}

		public static readonly Валюта Рубли = new Валюта(
			new ЕдиницаИзмерения(РодЧисло.Мужской, "рубль", "рубля", "рублей"),
			new ЕдиницаИзмерения(РодЧисло.Женский, "копейка", "копейки", "копеек"));

		public static readonly Валюта Доллары = new Валюта(
			new ЕдиницаИзмерения(РодЧисло.Мужской, "доллар США", "доллара США", "долларов США"),
			new ЕдиницаИзмерения(РодЧисло.Мужской, "цент", "цента", "центов"));

		public static readonly Валюта Евро = new Валюта(
			new ЕдиницаИзмерения(РодЧисло.Мужской, "евро", "евро", "евро"),
			new ЕдиницаИзмерения(РодЧисло.Мужской, "цент", "цента", "центов"));

		/// <summary>
		/// Возвращает пропись суммы строчными буквами.
		/// </summary>
		public string Пропись(decimal сумма) {
			return Сумма.Пропись(сумма, this);
		}

		/// <summary>
		/// Возвращает пропись суммы строчными буквами.
		/// </summary>
		public string Пропись(double сумма) {
			return Сумма.Пропись(сумма, this);
		}

		/// <summary>
		/// Возвращает пропись суммы.
		/// </summary>
		public string Пропись(decimal сумма, Заглавные заглавные) {
			return Сумма.Пропись(сумма, this, заглавные);
		}

		/// <summary>
		/// Возвращает пропись суммы.
		/// </summary>
		public string Пропись(double сумма, Заглавные заглавные) {
			return Сумма.Пропись(сумма, this, заглавные);
		}
	}

	/// <summary>
	/// Класс, хранящий падежные формы единицы измерения в явном виде.
	/// </summary>
	public class ЕдиницаИзмерения : IЕдиницаИзмерения {
		/// <summary> </summary>
		public ЕдиницаИзмерения(
			РодЧисло родЧисло,
			string именЕдин,
			string родЕдин,
			string родМнож) {
			this.родЧисло = родЧисло;
			this.именЕдин = именЕдин;
			this.родЕдин = родЕдин;
			this.родМнож = родМнож;
		}

		readonly РодЧисло родЧисло;
		readonly string именЕдин;
		readonly string родЕдин;
		readonly string родМнож;

		#region IЕдиницаИзмерения Members

		string IЕдиницаИзмерения.ИменЕдин {
			get { return this.именЕдин; }
		}

		string IЕдиницаИзмерения.РодЕдин {
			get { return this.родЕдин; }
		}

		string IЕдиницаИзмерения.РодМнож {
			get { return this.родМнож; }
		}

		РодЧисло IЕдиницаИзмерения.РодЧисло {
			get { return this.родЧисло; }
		}

		#endregion
	}

	#region РодЧисло

	/// <summary>
	/// Указывает род и число.
	/// Может передаваться в качестве параметра "единица измерения" метода 
	/// <see cref="Число.Пропись(decimal,IЕдиницаИзмерения,StringBuilder)"/>.
	/// Управляет родом и числом числительных один и два.
	/// </summary>
	/// <example>
	/// Число.Пропись (2, РодЧисло.Мужской); // "два"
	/// Число.Пропись (2, РодЧисло.Женский); // "две"
	/// Число.Пропись (21, РодЧисло.Средний); // "двадцать одно"
	/// </example>
	public abstract class РодЧисло : IЕдиницаИзмерения {
		internal abstract string ПолучитьФорму(IИзменяетсяПоРодам слово);

		#region Рода

		class _Мужской : РодЧисло {
			internal override string ПолучитьФорму(IИзменяетсяПоРодам слово) {
				return слово.Мужской;
			}
		}

		class _Женский : РодЧисло {
			internal override string ПолучитьФорму(IИзменяетсяПоРодам слово) {
				return слово.Женский;
			}
		}

		class _Средний : РодЧисло {
			internal override string ПолучитьФорму(IИзменяетсяПоРодам слово) {
				return слово.Средний;
			}
		}

		class _Множественное : РодЧисло {
			internal override string ПолучитьФорму(IИзменяетсяПоРодам слово) {
				return слово.Множественное;
			}
		}

		public static readonly РодЧисло Мужской = new _Мужской();
		public static readonly РодЧисло Женский = new _Женский();
		public static readonly РодЧисло Средний = new _Средний();
		public static readonly РодЧисло Множественное = new _Множественное();

		#endregion

		#region IЕдиницаИзмерения Members

		РодЧисло IЕдиницаИзмерения.РодЧисло {
			get { return this; }
		}

		string IЕдиницаИзмерения.ИменЕдин {
			get { return null; }
		}

		string IЕдиницаИзмерения.РодЕдин {
			get { return null; }
		}

		string IЕдиницаИзмерения.РодМнож {
			get { return null; }
		}

		#endregion
	}

	#region IИзменяетсяПоРодам

	internal interface IИзменяетсяПоРодам {
		string Мужской { get; }
		string Женский { get; }
		string Средний { get; }
		string Множественное { get; }
	}

	#endregion

	#endregion

	#region ЕдиницаИзмерения

	/// <summary>
	/// Представляет единицу измерения (например, метр, рубль)
	/// и содержит всю необходимую информацию для согласования
	/// этой единицы с числом, а именно - три падежно-числовых формы
	/// и грамматический род / число.
	/// </summary>
	public interface IЕдиницаИзмерения {
		/// <summary>
		/// Форма именительного падежа единственного числа.
		/// Согласуется с числительным "один":
		/// одна тысяча, один миллион, один рубль, одни сутки и т.д.
		/// </summary>
		string ИменЕдин { get; }

		/// <summary>
		/// Форма родительного падежа единственного числа.
		/// Согласуется с числительными "один, два, три, четыре":
		/// две тысячи, два миллиона, два рубля, двое суток и т.д.
		/// </summary>
		string РодЕдин { get; }

		/// <summary>
		/// Форма родительного падежа множественного числа.
		/// Согласуется с числительным "ноль, пять, шесть, семь" и др:
		/// пять тысяч, пять миллионов, пять рублей, пять суток и т.д.
		/// </summary>
		string РодМнож { get; }

		/// <summary>
		/// Род и число единицы измерения.
		/// </summary>
		РодЧисло РодЧисло { get; }
	}

	#endregion
}
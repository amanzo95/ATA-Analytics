using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MRP_Analyzer.Model
{
	public class NullToZeroConverter : JsonConverter<double>
	{
		public override double ReadJson(JsonReader reader, Type objectType, double existingValue, bool hasExistingValue, JsonSerializer serializer)
		{
			if (reader.TokenType == JsonToken.Null || reader.TokenType == JsonToken.String && string.IsNullOrEmpty((string)reader.Value))
			{
				return 0;
			}

			try
			{
				return Convert.ToDouble(reader.Value);
			}
			catch (Exception)
			{
				return 0;
			}
		}

		public override void WriteJson(JsonWriter writer, double value, JsonSerializer serializer)
		{
			serializer.Serialize(writer, value);
		}
	}

	public class MRPComparisionModel
	{
		public string Item { get; set; }
		public string Type { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Jan { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Feb { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Mar { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Apr { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double May { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Jun { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Jul { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Aug { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Sep { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Oct { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Nov { get; set; }
		[JsonConverter(typeof(NullToZeroConverter))]
		public double Dec { get; set; }
	}

	public class Month
	{
		public int Number { get; set; }
		public string Name { get; set; }

		public Month(int number, string name)
		{
			Number = number;
			Name = name;
		}
	}

	public static class MonthsList
	{
		public static readonly List<Month> monthsList;

		static MonthsList()
		{
			monthsList = new List<Month>
		{
			new Month(1, "Jan"),
			new Month(2, "Feb"),
			new Month(3, "Mar"),
			new Month(4, "Apr"),
			new Month(5, "May"),
			new Month(6, "Jun"),
			new Month(7, "Jul"),
			new Month(8, "Aug"),
			new Month(9, "Sep"),
			new Month(10, "Oct"),
			new Month(11, "Nov"),
			new Month(12, "Dec")
		};
		}
	}

	public class MRPFinalComparisionModel
	{
		public string Item { get; set; }
		public string Type { get; set; }
		public double Jan1 { get; set; } = 0;
		public double Jan2 { get; set; } = 0;
		public double JanDif { get { return Jan1 - Jan2; } }
		public double Feb1 { get; set; } = 0;
		public double Feb2 { get; set; } = 0;
		public double FebDif { get { return Feb1 - Feb2; } }
		public double Mar1 { get; set; } = 0;
		public double Mar2 { get; set; } = 0;
		public double MarDif { get { return Mar1 - Mar2; } }
		public double Apr1 { get; set; } = 0;
		public double Apr2 { get; set; } = 0;
		public double AprDif { get { return Apr1 - Apr2; } }
		public double May1 { get; set; } = 0;
		public double May2 { get; set; } = 0;
		public double MayDif { get { return May1 - May2; } }
		public double Jun1 { get; set; } = 0;
		public double Jun2 { get; set; } = 0;
		public double JunDif { get { return Jun1 - Jun2; } }
		public double Jul1 { get; set; } = 0;
		public double Jul2 { get; set; } = 0;
		public double JulDif { get { return Jul1 - Jul2; } }
		public double Aug1 { get; set; } = 0;
		public double Aug2 { get; set; } = 0;
		public double AugDif { get { return Aug1 - Aug2; } }
		public double Sep1 { get; set; } = 0;
		public double Sep2 { get; set; } = 0;
		public double SepDif { get { return Sep1 - Sep2; } }
		public double Oct1 { get; set; } = 0;
		public double Oct2 { get; set; } = 0;
		public double OctDif { get { return Oct1 - Oct2; } }
		public double Nov1 { get; set; } = 0;
		public double Nov2 { get; set; } = 0;
		public double NovDif { get { return Nov1 - Nov2; } }
		public double Dec1 { get; set; } = 0;
		public double Dec2 { get; set; } = 0;
		public double DecDif { get { return Dec1 - Dec2; } }
		public double TotalDif
		{
			get
			{
				return JanDif +
					   FebDif +
					   MarDif +
					   AprDif +
					   MayDif +
					   JunDif +
					   JulDif +
					   AugDif +
					   SepDif +
					   OctDif +
					   NovDif +
					   DecDif;
			}
		}
	}
}

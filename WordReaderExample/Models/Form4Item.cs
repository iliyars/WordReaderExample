using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WordReaderExample.Models
{
    public class Form4Item : DomainObject
    {
        public Form4Item()
        {
        }

        public Form4Item(string family, string name, int? quantity) : base(family, name, quantity)
        {
        }

        /// <summary>
        /// Наличие в перечнях при утверждении ТТЗ.
        /// </summary>
        public string? InListTTZ { get; set; }
        /// <summary>
        /// Наличие в перечнях последних редакций.
        /// </summary>
        public string? LastEditions { get; set; }
        /// <summary>
        /// Показатель ресурса, ч.
        /// </summary>
        public int? ResourceHours { get; set; }
        /// <summary>
        /// Показатель срока службы, лет.
        /// </summary>
        public int? LifeTimeYears { get; set; }
        /// <summary>
        /// показатель сохраняемости, лет.
        /// </summary>
        public int? PreservationYears { get; set; }
        /// <summary>
        /// Диапазон частот, Гц
        /// </summary>
        public string? FrequencyRange { get; set; }
        /// <summary>
        /// Уровень звукового давлени, дБ.
        /// </summary>
        public int? SoundPressure { get; set; }
        /// <summary>
        /// Линейное ускорение, М.С.Е-2(G).
        /// </summary>
        public string? LineAcceleration { get; set; }
        /// <summary>
        /// Давление окр.среды пониженное.
        /// </summary>
        public string? LowPressure { get; set; }
        /// <summary>
        /// Давление окр.среды повышенное.
        /// </summary>
        public string? HighPressure { get; set; }
        /// <summary>
        /// Предельная температура пониженная.
        /// </summary>
        public string? LowTemperature { get; set; }
        /// <summary>
        /// Предельная температура повышенная.
        /// </summary>
        public string? HighTemperature { get; set; }
        /// <summary>
        /// Относительня влажность, %.
        /// </summary>
        public int? HumidityPercent { get; set; }
        /// <summary>
        /// Относительня влажность, С.
        /// </summary>
        public string? HumidityCelcius { get; set; }
        /// <summary>
        /// Роса, иней.
        /// </summary>
        public string? Dew { get; set; }
        /// <summary>
        /// Стойкость к ВССФ.
        /// </summary>
        /// 
        public string? SpecialFactors { get; set; }
        public string AdditionalFormName { get; set; }

        public override List<string> GetParamsForWord()
        {
            return new List<string>
            {
                Type,
                Name,
                Quantity?.ToString() ?? "\u2014",
                InListTTZ ?? "\u2014",
                LastEditions ?? "\u2014",
                ResourceHours?.ToString() ?? "\u2014",
                LifeTimeYears?.ToString() ?? "\u2014",
                PreservationYears?.ToString() ?? "\u2014",
                FrequencyRange ?? "\u2014",
                SoundPressure?.ToString() ?? "\u2014",
                LineAcceleration?.ToString() ?? "\u2014",
                LowPressure ?.ToString() ?? "\u2014",
                HighPressure?.ToString() ?? "\u2014",
                LowTemperature ?.ToString() ?? "\u2014",
                HighTemperature ?.ToString() ?? "\u2014",
                HumidityPercent ?.ToString() ?? "\u2014",
                HumidityCelcius ?? "\u2014",
                SpecialFactors ?.ToString() ?? "\u2014",
                Note ?? "\u2014"
            };
        }
    }
}

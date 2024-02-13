using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WordReaderExample.Models
{
    public class Form68Item : DomainObject
    {
        /// <summary>
        /// Постояное напряжение.
        /// </summary>
        public string? DcVoltage { get; set; }
        /// <summary>
        /// Переменное напряжение.
        /// </summary>
        public string? AcVoltage { get; set; }
        /// <summary>
        /// Импульсное напряжение.
        /// </summary>
        public string? ImpulseVoltage { get; set; }
        /// <summary>
        /// Суммарное напряжение.
        /// </summary>
        public string? SumVoltage { get; set; }
        /// <summary>
        /// Частота, Гц
        /// </summary>
        public string? Frequancy { get; set; }
        /// <summary>
        /// Длительность импульса.
        /// </summary>
        public string? ImpulseDuration { get; set; }
        /// <summary>
        /// Импульсная мощность.
        /// </summary>
        public string? ImpulsePower { get; set; }
        /// <summary>
        /// Средняя мощность.
        /// </summary>
        public string? MeanPower { get; set; }
        /// <summary>
        /// Коэффициент нагрузки (импульсный режим).
        /// </summary>
        public string? LoadKoeffImpulse { get; set; }
        /// <summary>
        /// Ток через подвижный контакт.
        /// </summary>
        public string? CurrentMovingContact { get; set; }

        /// <summary>
        /// Температура окружающей среды.
        /// </summary>
        public string? AmbientTemperature { get; set; }
        /// <summary>
        /// Температура перегрева.
        /// </summary>
        public string? SuperHeatTemperature { get; set; }
        /// <summary>
        /// Суммарная мощность.
        /// </summary>
        public string? SumPower { get; set; }
        /// <summary>
        /// Температура окружающей среды(корпуса).
        /// </summary>
        public string? AmbientTemperatureCase { get; set; }
        /// <summary>
        /// Коэффициент нагрузки.
        /// </summary>
        public string? LoadKoeff { get; set; }



        public override List<string> GetParamsForWord()
        {
            if (Note == string.Empty) Note = "\u2014";
            return new List<string>()
            {
                PositionNames,
                Name,
                DcVoltage?.ToString() ?? "\u2014",
                AcVoltage?.ToString() ?? "\u2014",
                ImpulseVoltage?.ToString() ?? "\u2014",
                SumVoltage?.ToString() ?? "\u2014",
                Frequancy?.ToString() ?? "\u2014",
                ImpulseDuration?.ToString() ?? "\u2014",
                ImpulsePower?.ToString() ?? "\u2014",
                MeanPower?.ToString() ?? "\u2014",
                LoadKoeffImpulse?.ToString() ?? "\u2014",
                CurrentMovingContact?.ToString() ?? "\u2014",
                AmbientTemperature?.ToString() ?? "\u2014",
                SuperHeatTemperature?.ToString() ?? "\u2014",
                SumPower?.ToString() ?? "\u2014",
                AmbientTemperature?.ToString() ?? "\u2014",
                LoadKoeff ?? "\u2014",
                Note?.ToString() ?? "\u2014",
            };
        }
    }
}

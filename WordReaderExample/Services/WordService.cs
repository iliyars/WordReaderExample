using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static System.Collections.Specialized.BitVector32;
using static System.Net.Mime.MediaTypeNames;
using System.Xml.Linq;
using WordReaderExample.Models.WordDocuments;
using WordReaderExample.Models;
using Word = Microsoft.Office.Interop.Word;


namespace WordReaderExample.Services
{

    public delegate void Read(ref int row, ref int collIndex, ref int collNumAddition, ref int collHeaderAddition, Word.Table table, DomainObject item, WordDocument wordDocument);

    public class WordService
    {
        private object missing = Missing.Value;
        
        public List<DomainObject> ReadWordDocument(string filePath)
        {
            List<string> tableHeaders = new List<string>();
            Word.Application wordApp = null;
            Word.Document wordDocument = null;
            List<DomainObject> allItems = new List<DomainObject>();
            try
            {
                wordApp = new Word.Application();
                wordDocument = wordApp.Documents.Open(filePath);
                wordApp.Visible = false;

                for (int i = 1; i < wordDocument.Tables.Count; i++)
                {
                    Word.Table table = wordDocument.Tables[i];
                    string tableHeader = table.Cell(1, 1).Range.Text;
                    int index = tableHeader.IndexOf('\r');
                    if (index > 0)
                    {
                        tableHeader = tableHeader.Substring(0, index);
                    }
                    tableHeaders.Add(tableHeader.Trim());

                }
                for (int i = 0; i < tableHeaders.Count; i++)
                {
                    //var a = Enum.TryParse(tableHeaders[i], out FormName formName);
                    switch (tableHeaders[i])
                    {
                        case "ФОРМА 4":
                            var items = ReadForm4(wordDocument.Tables[i + 1]);
                            allItems.AddRange(items);
                            break;

                        case "ФОРМА 68":
                            allItems.AddRange(ReadForm68(wordDocument.Tables[i + 1]));
                            break;
                        default:
                            return null;
                    }
                }
                allItems = allItems.DistinctBy(x => new { x.Name, x.FormName }).ToList();

                return allItems;


            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                wordDocument.Close(Word.WdSaveOptions.wdSaveChanges);
                wordApp.Quit();
            }
        }
        public List<Form4Item> ReadForm4(Word.Table table)
        {
            List<Form4Item> form4Items = new List<Form4Item>();
            Form4Item form4Item = new Form4Item();
            Form4WordDocument wordDocument = new Form4WordDocument();
            int row = wordDocument.ParametersStartRow;
            int collIndex = 0;
            int collNumAddition = 0;
            for (int i = 1; i <= wordDocument.ItemColumnsOnPage; i++)
            {
                form4Item.Type = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form4Item.Name = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                if (form4Item.Name == " " || form4Item.Name == "")
                {
                    continue;
                }
                if (form4Items.FirstOrDefault(i => i.Name == form4Item.Name) != null) break;
                form4Item.Family = form4Item.Name;
                row++; collIndex++;
                form4Item.InListTTZ = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form4Item.LastEditions = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                int.TryParse(table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r'), out int resourceHours);
                form4Item.ResourceHours = resourceHours;
                int.TryParse(table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r'), out int lifeTimeYears);
                form4Item.LifeTimeYears = lifeTimeYears;
                int.TryParse(table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r'), out int preservationYears);
                form4Item.PreservationYears = preservationYears;
                form4Item.FrequencyRange = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                int.TryParse(table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r'), out int soundPressure);
                form4Item.SoundPressure = soundPressure;
                form4Item.LineAcceleration = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form4Item.LowPressure = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form4Item.HighPressure = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form4Item.LowTemperature = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form4Item.HighTemperature = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                int.TryParse(table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r'), out int humidityPercent);
                form4Item.HumidityPercent = humidityPercent;
                form4Item.HumidityCelcius = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form4Item.Dew = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form4Item.SpecialFactors = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form4Item.Note = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');

                form4Item.FormName = FormName.Form4.ToString();
                form4Items.Add((Form4Item)form4Item.Clone());
                row = wordDocument.ParametersStartRow;
                collNumAddition++;
                collIndex = 0;
            }
            return form4Items;
        }
        public List<Form68Item> ReadForm68(Word.Table table)
        {
            List<Form68Item> form68items = new List<Form68Item>();
            Form68Item form68item = new Form68Item();
            Form68WordDocument wordDocument = new Form68WordDocument();
            int row = wordDocument.ParametersStartRow;
            int collIndex = 0;
            int collNumAddition = 0;
            int collHeaderAddition = 0;
            for (int i = 1; i <= wordDocument.ItemColumnsOnPage; i++)
            {
                row++; collIndex++;
                form68item.Name = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collHeaderAddition).Range.Text.Trim('\a', '\r');
                if (form68item.Name == " " || form68item.Name == "" || form68items.FirstOrDefault(i => i.Name == form68item.Name) != null)
                {
                    continue;
                }

                row++;
                form68item.DcVoltage = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.AcVoltage = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.ImpulseVoltage = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.SumVoltage = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.Frequancy = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.ImpulseDuration = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.ImpulsePower = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.MeanPower = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.LoadKoeffImpulse = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.CurrentMovingContact = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.AmbientTemperature = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.SuperHeatTemperature = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.SumPower = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.AmbientTemperatureCase = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                form68item.LoadKoeff = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collNumAddition).Range.Text.Trim('\a', '\r');
                string output = Regex.Replace(form68item.LoadKoeff, @"\(.*?\)", "");
                form68item.LoadKoeff = output;
                string note = table.Cell(row++, wordDocument.ParametersColls[collIndex++] + collHeaderAddition).Range.Text.Trim('\a', '\r');
                form68item.Note = ParseNote(note);
                form68item.Type = "Резистор";
                form68item.FormName = FormName.Form68.ToString();
                form68items.Add((Form68Item)form68item.Clone());
                row = wordDocument.ParametersStartRow;
                collNumAddition += 2;
                collHeaderAddition++;
                collIndex = 0;
            }
            return form68items;
        }
        private string ParseNote(string str)
        {
            string note = "";
            if (!str.Contains('*'))
                return str;
            string[] notes = str.Split('*');
            foreach (string s in notes)
            {
                if (s != "")
                    note += s + "!";
            }
            return note;

        }
    }
}

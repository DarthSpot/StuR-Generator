using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using Nager.Date;

namespace StuR_Generator
{
    class Program
    {

        private static (int, int) JanFirst { get; } = (2, 2);
        private static (int, int) DaySumOne { get; } = (2, 41);

        static void Main(string[] args)
        {
            Console.WriteLine("Generating Excel for 2020");
            var wb = GetTemplateXlsx();
            var inputSheet = wb.Worksheets.Worksheet("Eingabe");
            var outputSheet = wb.Worksheets.Worksheet("Ausgabe");
            
            var cal = CreateCalendar(2020);

            DayColors(inputSheet, cal);
            DayColors(outputSheet, cal);
            FillDate(outputSheet, cal);


            wb.Save();

        }

        private static void FillDate(IXLWorksheet outputSheet, Dictionary<(int, int), DayType> cal)
        {
            for (var month = 0; month < 12; month++)
            {
                var field = outputSheet.Cell(DaySumOne.Item2, DaySumOne.Item1 + month);
                var daysOfWork = cal.Count(x => x.Key.Item2 == month+1 && x.Value == DayType.WorkDay);
                field.Value = daysOfWork;
            }
        }

        public enum DayType
        {
            None,
            Weekend,
            WorkDay,
            Holiday,
            SpecialHoliday
        }
        private static Dictionary<(int,int), DayType> CreateCalendar(int year)
        {
            var res = new Dictionary<(int,int), DayType>();

            // Normale Tage
            for (var date = DateTime.Parse($"01.01.{year}");
                date <= DateTime.Parse($"31.12.{year}");
                date = date.AddDays(1))
            {
                if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                    res.Add((date.Day, date.Month), DayType.Weekend);
                else
                    res.Add((date.Day, date.Month), DayType.WorkDay);
            }

            // Feiertage
            foreach (var holiday in DateSystem.GetPublicHoliday(DateTime.Now.Year, CountryCode.DE).Where(x => 
                x.Global ||
                x.Counties.Contains("DE-HH")))
            {
                res[(holiday.Date.Day, holiday.Date.Month)] = DayType.Holiday;
            }

            // Freie Zeit zwischen den Tagen
            for (var date = DateTime.Parse($"24.12.{year}");
                date <= DateTime.Parse($"31.12.{year}");
                date = date.AddDays(1))
            {
                if (res[(date.Day, date.Month)] == DayType.WorkDay)
                    res[(date.Day, date.Month)] = DayType.SpecialHoliday;
            }

            return res;
        }

        private static void DayColors(IXLWorksheet sheet, Dictionary<(int, int), DayType> cal)
        {
            for (var month = 0; month < 12; month++)
            {
                for (var day = 0; day < 31; day++)
                {
                    var fieldT = (month:JanFirst.Item1 + month, day:JanFirst.Item2 + day);
                    var field = sheet.Cell(fieldT.day, fieldT.month);
                    var date = (day+1, month+1);
                    XLColor color = null;
                    if (cal.ContainsKey(date))
                    {
                        switch (cal[date])
                        {
                            case DayType.Weekend:
                                color = XLColor.AirForceBlue;
                                break;
                            case DayType.WorkDay:
                                color = XLColor.White;
                                break;
                            case DayType.Holiday:
                                color = XLColor.Yellow;
                                break;
                            case DayType.SpecialHoliday:
                                color = XLColor.Orange;
                                break;
                        }
                    }
                    else
                    {
                        color = XLColor.Black;
                    }
                    field.Style.Fill.BackgroundColor = color;
                }
            }
        }

        private static XLWorkbook GetTemplateXlsx()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "StuR_Generator.Data.Template.xlsx";

            using var stream = assembly.GetManifestResourceStream(resourceName);
            using var mem = new MemoryStream();

            stream.CopyTo(mem);

            File.WriteAllBytes("Stundenrechner.xlsx", mem.ToArray());

            return new XLWorkbook("Stundenrechner.xlsx");
        }
    }
}

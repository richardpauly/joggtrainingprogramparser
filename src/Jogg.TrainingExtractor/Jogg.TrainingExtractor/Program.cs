using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Ical.Net;
using Ical.Net.DataTypes;
using Ical.Net.Interfaces.DataTypes;
using Ical.Net.Serialization;
using Ical.Net.Serialization.iCalendar.Serializers;

namespace Jogg.TrainingExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            const string url = "http://jogg.se/Program/Program.aspx?id=14&raceid=649167";
            var web = new HtmlWeb();
            var doc = web.Load(url);
            var header = doc.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[3]/div/div/h2")[0].InnerText.Replace(':','.');
            var rows = doc.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[3]/div/div/table")[0].Descendants("tr");
            var calendar = new Calendar();
            calendar.AddTimeZone(new VTimeZone("Europe/Copenhagen"));
            foreach (var row in rows.Where(row => row.Id.Contains("weekRepeater")))
            {
                var cells = row.Descendants("td").ToList();
                var date = ParseDateForRow(cells);
                var subject = ParseSubject(cells);
                var details = ParseDetails(cells);
                if (!subject.StartsWith("Vila"))
                {
                    calendar.Events.Add(new Event()
                    {

                        DtStart = new CalDateTime(date.Year, date.Month, date.Day),
                        DtEnd = new CalDateTime(date.Year, date.Month, date.Day),
                        Summary = subject,
                        Description = details,
                        Class = "PUBLIC",
                        
                    });
                }
            }
            var serializer = new CalendarSerializer(new SerializationContext());
            var serializedCalendar = serializer.SerializeToString(calendar);
            var directory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
            var filename = header + ".ics";
            var path = Path.Combine(directory,filename);
            using (var streamWriter = new StreamWriter(path))
            {
                streamWriter.Write(serializedCalendar);
                streamWriter.Close();
            }
        }

        private static string ParseSubject(IReadOnlyList<HtmlNode> cells)
        {
            return GetFirstWord(cells[1]) + " " + cells[2].InnerText.Trim();
        }

        private static string GetFirstWord(HtmlNode cell)
        {
            return cell.InnerText.Trim().Split(' ')[0];
        }

        private static string ParseDetails(IReadOnlyList<HtmlNode> cells)
        {
            var listItems = cells[1].Descendants("li").Select(li => li.InnerText.Trim());
            return string.Join("\r\r", listItems);
        }

        private static DateTime ParseDateForRow(IReadOnlyList<HtmlNode> cells)
        {
            var dateParts = cells[0].Descendants("span").First().InnerText.Trim().Split(' ');
            var day = int.Parse(dateParts[0]);
            var month = ParseMonth(dateParts[1]);
            return new DateTime(DateTime.Now.Year, month, day);
        }

        private static int ParseMonth(string month)
        {
            switch (month.ToLower())
            {
                case "januari":
                    return 1;
                case "februari":
                    return 2;
                case "mars":
                    return 3;
                case "april":
                    return 4;
                case "maj":
                    return 5;
                case "juni":
                    return 6;
                case "juli":
                    return 7;
                case "augusti":
                    return 8;
                case "september":
                    return 9;
                case "oktober":
                    return 10;
                case "november":
                    return 11;
                case "december":
                    return 12;
                default:
                    throw new ArgumentOutOfRangeException(nameof(month), month);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HtmlAgilityPack;
using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;

namespace Jogg.TrainingExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Du måste ange en url där ett träningsprogram kan hämtas");
                Console.ReadKey();
                return;
            }
            var url = args[0];
            Console.WriteLine("Försöker läsa träningsprogram från {0}", url);
            var web = new HtmlWeb();
            var doc = web.Load(url);
            var header = doc.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div/div/h2")[0].InnerText.Replace(':', '.');
            var raceDate = ParseDateWithYear(doc.DocumentNode.SelectSingleNode("//*[@id=\"MainContent_m_adaptionDiv\"]/b/text()[2]").InnerText.Replace("-", "").Replace("&nbsp;", "").Trim());
            var minimumDate = raceDate.AddMonths(-6); // Expecting no programs to start more than 6 months before competition
            var competition = ParseCompetitionName(doc, header);
            var rows = doc.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div/div/table")[0].Descendants("tr");
            var calendar = new Calendar();
            Console.WriteLine("Exporterar kalender från träningsprogrammet...");
            calendar.AddTimeZone(new VTimeZone("Europe/Copenhagen"));
            foreach (var row in rows.Where(row => row.Id.Contains("weekRepeater") && !row.InnerHtml.Contains("MidPanorama")))
            {
                var cells = row.Descendants("td").ToList();
                var subject = ParseSubject(cells);
                if (!subject.StartsWith("Vila"))
                {
                    var date = ParseDateForRow(cells, minimumDate);
                    var duration = ParseDuration(cells);
                    var start = new DateTime(date.Year, date.Month, date.Day, 12, 0, 0);
                    var end = start.Add(duration);
                    var link = url + "#" + row.Id;
                    var details = ParseDetails(cells, link) + $"\n\nMål: {competition}";

                    calendar.Events.Add(new CalendarEvent()
                    {

                        DtStart = new CalDateTime(start),
                        DtEnd = new CalDateTime(end),
                        Summary = subject,
                        Description = details,
                        Class = "PUBLIC",
                        IsAllDay = false,
                        Categories = new List<string>() { "Träning" }
                    });
                }
            }
            var serializer = new CalendarSerializer(new SerializationContext());
            var serializedCalendar = serializer.SerializeToString(calendar);
            var directory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
            var filename = $"{competition} - {header}.ics";
            var path = Path.Combine(directory, filename);
            using (var streamWriter = new StreamWriter(path))
            {
                streamWriter.Write(serializedCalendar);
                streamWriter.Close();
            }
            Console.WriteLine("Klart! Träningsprogrammet sparades till {0}", path);
        }

        private static string ParseCompetitionName(HtmlDocument doc, string fallback)
        {
            string competition;
            try
            {
                // If the competition is named
                competition = doc.DocumentNode.SelectNodes("//*[@id=\"MainContent_m_competitionlink\"]/b")[0].InnerText;

            }
            catch (Exception)
            {
                // If the competition is not named, but rather a date
                competition = doc.DocumentNode.SelectNodes("//*[@id=\"MainContent_m_adaptionDiv\"]/text()[2]")[0].InnerText.Replace("&nbsp;", "").Trim().FirstCharToUpper();
                if (!competition.StartsWith("Tävling"))
                    competition = fallback;
            }

            return competition;
        }

        private static string ParseSubject(IReadOnlyList<HtmlNode> cells)
        {
            return GetFirstWord(cells[1]) + " " + cells[2].InnerText.Trim();
        }

        /// <summary>
        /// Get the length in kilometers for the given row
        /// </summary>
        private static int ParseLength(IReadOnlyList<HtmlNode> cells, int defaultValue = 10)
        {
            var parts = cells[2].InnerText?.Replace("km", String.Empty).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(part =>
            {
                if (int.TryParse(part, out int value))
                    return value;
                return default;
            }).Where(part => part != default);
            return (parts != null && parts.Any()) ? parts.Max() : defaultValue;
        }

        private static TimeSpan ParseDuration(IReadOnlyList<HtmlNode> cells)
        {
            var length = ParseLength(cells);
            return TimeSpan.FromHours(length / 10d * 1.15); // one hour per 10k, and add some extra percentage
        }

        private static string GetFirstWord(HtmlNode cell)
        {
            return cell.InnerText.Trim().Split(' ')[0];
        }

        private static string ParseDetails(IReadOnlyList<HtmlNode> cells, string link)
        {
            var listItems = cells[1].Descendants("li").Select(li => li.InnerText.Trim()).Where(item => !item.StartsWith("Tips"));
            return string.Join("\r\r", listItems) + $"\r\nDetaljer: {link}";
        }

        private static DateTime ParseDateForRow(IReadOnlyList<HtmlNode> cells, DateTime minimumDate)
        {
            var dateParts = cells[0].Descendants("span").First().InnerText.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var day = int.Parse(dateParts[0]);
            var month = ParseMonth(dateParts[1]);
            var date = new DateTime(DateTime.Now.Year, month, day);

            while (date < minimumDate)
                date = date.AddYears(1);
            return date;
        }
        private static DateTime ParseDateWithYear(string input)
        {
            if (input.StartsWith("tävling"))
            {
                input = input.Replace("tävling", "").Trim();
                var day = int.Parse(input.Substring(6, 2));
                var month = int.Parse(input.Substring(4, 2));
                var year = int.Parse(input.Substring(0, 4));
                return new DateTime(year, month, day);
            }
            else
            {
                var dateParts = input.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                var day = int.Parse(dateParts[0]);
                var month = ParseMonth(dateParts[1]);
                var year = int.Parse(dateParts[2]);
                return new DateTime(year, month, day);
            }
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

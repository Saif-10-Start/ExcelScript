using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace ExcelScript
{
    public static class Extensions
    {
        public static string GetName(this HeaderName header)
            => header switch
            {
                HeaderName.Item => "item",
                HeaderName.Supplier => "supplier",
                HeaderName.Email => "email",
                HeaderName.PhoneNumber => "phone number",
                HeaderName.Location => "location",
                HeaderName.Website => "website",
                _ => throw new ArgumentException($"Unknown header: {header}")
            };
        public static HeaderName GetHeader(string val)
        {
            var normalized = val?.ToLower().Trim() ?? string.Empty;
            return normalized switch
            {
                "item" => HeaderName.Item,
                "supplier" => HeaderName.Supplier,
                "name" => HeaderName.Supplier,
                "email" => HeaderName.Email,
                "phone number" => HeaderName.PhoneNumber,
                "location" => HeaderName.Location,
                "website" => HeaderName.Website,
                _ => throw new ArgumentException($"Unknown header: {val}")
            };
        }
        public static HeaderName[] GetHeaders()
            => Enum.GetValues<HeaderName>();
        public static IEnumerable<string> GetHeaderNames()
        {
            var values = GetHeaders();
            return values.Select(val => val.GetName());
        }

        public static bool IsMainColumn(this IXLColumn column)
        {
            if (column.CellsUsed().Count() <= 0) throw new Exception("No used cell was found!");
            var cell = column.Cell(1);
            try
            {
                GetHeader(cell.Value.ToString());
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        public static HeaderName GetHeader(this IXLColumn column)
            => GetHeader(column.FirstCellUsed().Value.ToString());

        public static void CleanLink(this IXLCell cell)
        {
            var url = cell.GetString();
            if (string.IsNullOrWhiteSpace(url)) return;

            var cleaned = url.Trim();

            cleaned = Regex.Replace(cleaned, @"^https?://", "");
            cleaned = Regex.Replace(cleaned, @"^www\.", "");

            cell.SetHyperlink(new XLHyperlink($"https://{cleaned}", $"Visit {cleaned}"));
            cell.Value = cleaned;
        }

        public static void FormatPhoneNumber(this IXLCell cell)
        {
            var number = cell.Value.ToString().Replace(" ", "");
            number = Regex.Replace(number, @"^\+?365", "");

            number = number.Insert(3, " ");
            number = "+365 " + number;

            cell.SetHyperlink(new XLHyperlink($"CALLTO:{number}", $"Call {number}"));
            cell.Value = number;
        }

        public static void FormatMail(this IXLCell cell)
        {
            var email = cell.Value.ToString().Replace(" ", "");

            cell.SetHyperlink(new XLHyperlink($"MAILTO:{email}", $"Send a mail to {email}"));
            cell.Value = email;
        }



        public static string FirstCharToUpper(this string input) =>
        input switch
        {
            null => throw new ArgumentNullException(nameof(input)),
            "" => throw new ArgumentException($"{nameof(input)} cannot be empty", nameof(input)),
            _ => input[0].ToString().ToUpper() + input.Substring(1)
        };
    }
}

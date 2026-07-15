using ClosedXML.Excel;
using System.Text.RegularExpressions;

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
                _ => HeaderName.Unknown
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

        public static IXLCell Strip(this IXLCell cell)
        {
            var value = cell.Value.ToString().Trim();
            value = Regex.Replace(value, @"(?i)^n[./\-]?a$", "");
            value = Regex.Replace(value, @"^[\-|_|—]+", "");

            cell.Value = value;
            return cell;
        }

        public static IXLCell CleanLink(this IXLCell cell)
        {
            var url = cell.GetString();
            if (cell.HasHyperlink) url = cell.GetHyperlink().ExternalAddress.ToString();
            if (string.IsNullOrWhiteSpace(url)) return cell;

            var cleaned = url.Trim();

            cleaned = Regex.Replace(cleaned, @"^https?://", "");
            cleaned = Regex.Replace(cleaned, @"^www\.", "");

            cell.SetHyperlink(new XLHyperlink($"https://{cleaned}", $"Visit {cleaned}"));
            cell.Value = cleaned;
            return cell;
        }

        public static IXLCell FormatPhoneNumber(this IXLCell cell)
        {
            var number = cell.Value.ToString().Replace(" ", "").Trim('/');
            if (string.IsNullOrEmpty(number)) return cell;

            var numbers = number.Split('/');
            List<string> formattedNumbers = [];

            foreach (var num in numbers)
            {
                var formatted = format(num);
                formattedNumbers.Add(formatted);
            }

            return cell;

            static string format(string num)
            {
                if (string.IsNullOrWhiteSpace(num)) return string.Empty;

                var cleaned = Regex.Replace(num, @"[^\d+]", "");
                var nationalNumber = Regex.Replace(cleaned, @"^(\+?356|\+?365|00356|00365|0+)", "");

                if (nationalNumber.Length < 4)
                    return "+356 " + nationalNumber;

                var formattedNational = nationalNumber.Insert(4, " ");
                return "+356 " + formattedNational;
            }
        }

        public static IXLCell FormatMail(this IXLCell cell)
        {
            var email = cell.Value.ToString().Replace(" ", "");
            if (string.IsNullOrEmpty(email)) return cell;
            if (!Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$")) return cell;

            cell.SetHyperlink(new XLHyperlink($"MAILTO:{email}", $"Send a mail to {email}"));
            cell.Value = email;
            return cell;
        }
        
        public static IXLCell FormatLocation(this IXLCell cell)
        {
            var location = cell.Value.ToString().Trim();
            if (string.IsNullOrEmpty(location)) return cell;

            var encodedLocation = Uri.EscapeDataString(location);

            cell.SetHyperlink(new XLHyperlink($"https://www.google.com/maps/search/?api=1&query={encodedLocation}", $"Search {location} on Google Maps"));
            cell.Value = location;
            return cell;
        }

        public static bool IsComment(this IXLCell cell)
        {
            var value = cell.Value.ToString().Trim();
            if (string.IsNullOrEmpty(value)) return false;

            return value.StartsWith("#");
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

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

using databaseAPI;

using GNA_CommercialLicenseValidator;

using gnaDataClasses;

using GNAgeneraltools;

using GNAspreadsheettools;

using OfficeOpenXml;

using T4Dlibrary;

using Twilio.TwiML.Voice;
using Twilio.Types;

using static GNAgeneraltools.gnaTools;

#pragma warning disable CS0219
#pragma warning disable CS8321
#pragma warning disable CS8600
#pragma warning disable CS8604



namespace GNA_exportTiltData
{






    internal class Program
    {
        static void Main(string[] args)
        {

            try
            {

                #region Setting state
                Console.OutputEncoding = System.Text.Encoding.Unicode;
                //if (Environment.UserInteractive && !Console.IsOutputRedirected) Console.Clear();
                int headingNo = 1;
                const string strTab1 = "     ";
                const string strTab2 = "        ";
                const string strTab3 = "           ";
                int exitCode = 0;
                string strFreezeScreen = "Yes";
                #endregion

                #region Instantiate core classes
                gnaTools gnaT = new();
                dbAPI gnaDBAPI = new();

                spreadsheetAPI gnaSpreadsheetAPI = new spreadsheetAPI(db: gnaDBAPI);
                string strSystemLogsFolder = @"C:\__SystemLogs\";
                Console.OutputEncoding = System.Text.Encoding.Unicode;
                T4Dapi t4dapi = new();
                t4dapi.SetCommercial("Dm4eGwoTaGxqY2hv");
                #endregion


                #region Header
                gnaT.WelcomeMessage($"GNA_exportTiltData {BuildInfo.BuildDateString()}");
                #endregion

                #region Config validation
                Console.WriteLine($"{headingNo++}. System Check");
                Console.Out.Flush();
                gnaT.VerifyLocalConfig();
                Console.WriteLine($"{strTab1}VerifyLocalConfig returned OK");
                Console.Out.Flush();
                #endregion

                #region Read config early
                NameValueCollection config = ConfigurationManager.AppSettings;
                strFreezeScreen = CleanConfig(config["freezeScreen"]);
                if (strFreezeScreen.Length == 0) strFreezeScreen = "Yes";
                #endregion


                #region License validation
                Console.WriteLine($"{headingNo++}. Validating the software license");
                string licenseCode = CleanConfig(config["LicenseCode"]);
                if (licenseCode.Length == 0)
                    throw new ConfigurationErrorsException("LicenseCode missing/empty.");
                LicenseValidator.ValidateLicense("TLTEXP", licenseCode);
                Console.WriteLine($"{strTab1}Validated");
                #endregion

                #region EPPlus license
                gnaT.epplusLicense();
                #endregion












                #region Clean exit
                void FinishAndExit()
                {
                    Console.WriteLine("\nGNA_exportTiltData completed...\n\n");
                    gnaT.freezeScreen(strFreezeScreen);
                }
                #endregion
ThatsAllFolks:
                FinishAndExit();
            }
            catch (Exception ex)
            {
                File.WriteAllText("fatal_crash.log", ex.ToString());
            }

        }


        #region Config helpers
        static string CleanConfig(string s) => (s ?? string.Empty).Trim().Trim('\'', '"');

        static string GetRequired(NameValueCollection cfg, string key)
        {
            string v = CleanConfig(cfg[key]);
            if (v.Length == 0)
                throw new ConfigurationErrorsException($"Missing/empty config key '{key}'.");
            return v;
        }

        static int GetRequiredInt(NameValueCollection cfg, string key, int minValueInclusive = int.MinValue, int maxValueInclusive = int.MaxValue)
        {
            string s = GetRequired(cfg, key);
            if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v))
                throw new ConfigurationErrorsException($"Config key '{key}' is invalid (expected integer). Value='{s}'.");
            if (v < minValueInclusive || v > maxValueInclusive)
                throw new ConfigurationErrorsException($"Config key '{key}' is out of range. Value={v}.");
            return v;
        }

        static bool IsYes(string s) => string.Equals(CleanConfig(s), "Yes", StringComparison.OrdinalIgnoreCase);
        #endregion

        #region General helpers
        private static string BuildManualEmailTime(string manualBlockEnd)
        {
            if (string.IsNullOrWhiteSpace(manualBlockEnd))
                return string.Empty;

            string tmp = manualBlockEnd.Replace("-", "")
                                       .Replace(" ", "_")
                                       .Replace(":", "h") + "m";
            return tmp.Length >= 14 ? tmp.Substring(0, 14) : tmp;
        }




        private static string SanitizeForFilename(string s)
        {
            s = (s ?? string.Empty).Trim();
            if (s.Length == 0) return "empty";

            foreach (char c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');

            s = s.Replace(" ", "_").Replace(":", "-");
            return s;
        }

        private static string FormatUtcForFilename(string utcTimestamp)
        {
            if (!DateTime.TryParseExact(
                    utcTimestamp,
                    new[] { "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm" },
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out DateTime dt))
            {
                throw new FormatException($"Invalid UTC timestamp format: '{utcTimestamp}'");
            }

            return dt.ToString("yyyyMMdd_HHmm", CultureInfo.InvariantCulture)
                     .Insert(11, "h");
        }




            public static string BuildDateString()
            {
                var buildDate =
                    System.IO.File.GetLastWriteTime(
                        Assembly.GetExecutingAssembly().Location);
                return buildDate.ToString("yyyy-MM-dd HH:mm");
            }



        #endregion


        #region Internal helpers
        internal static class ConfigParsing
        {
            public static bool GetBoolYesNo(System.Collections.Specialized.NameValueCollection appSettings, string key)
            {
                string? raw = appSettings[key];
                if (raw is null)
                    throw new ConfigurationErrorsException($"Missing required appSetting: '{key}'.");

                string value = raw.Trim();

                if (value.Equals("Yes", StringComparison.OrdinalIgnoreCase)) return true;
                if (value.Equals("No", StringComparison.OrdinalIgnoreCase)) return false;

                throw new ConfigurationErrorsException(
                    $"Invalid value for appSetting '{key}': '{raw}'. Expected 'Yes' or 'No'.");
            }

            public static string GetRequiredString(System.Collections.Specialized.NameValueCollection appSettings, string key)
            {
                string? raw = appSettings[key];
                if (string.IsNullOrWhiteSpace(raw))
                    throw new ConfigurationErrorsException($"Missing or empty required appSetting: '{key}'.");
                return raw.Trim();
            }

            public static int GetRequiredInt(System.Collections.Specialized.NameValueCollection appSettings, string key)
            {
                string raw = GetRequiredString(appSettings, key);
                if (!int.TryParse(raw, out int value))
                    throw new ConfigurationErrorsException($"Invalid integer for appSetting '{key}': '{raw}'.");
                return value;
            }

            public static double GetRequiredDouble(System.Collections.Specialized.NameValueCollection appSettings, string key)
            {
                string raw = GetRequiredString(appSettings, key);
                if (!double.TryParse(
                        raw,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out double value))
                {
                    throw new ConfigurationErrorsException(
                        $"Invalid double for appSetting '{key}': '{raw}'. Use '.' as decimal separator.");
                }
                return value;
            }
        }


        #endregion


    }
}
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.PortableExecutable;
using System.Security.AccessControl;

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
            #region Instantiate core classes
            gnaTools gnaT = new();
            dbAPI gnaDBAPI = new();
            spreadsheetAPI gnaSpreadsheetAPI = new spreadsheetAPI(db: gnaDBAPI);
            string strSystemLogsFolder = @"C:\__SystemLogs\";
            Console.OutputEncoding = System.Text.Encoding.Unicode;
            T4Dapi t4dapi = new();
            t4dapi.SetCommercial("Dm4eGwoTaGxqY2hv");
            int exitCode = 1;
            #endregion


            try
            {

                #region Setting state
                Console.OutputEncoding = System.Text.Encoding.Unicode;
                //if (Environment.UserInteractive && !Console.IsOutputRedirected) Console.Clear();
                int headingNo = 1;
                const string strTab1 = "     ";
                const string strTab2 = "        ";
                const string strTab3 = "           ";
                string strFreezeScreen = "Yes";
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
                bool freezeScreen = ConfigParsing.GetBoolYesNo(config, "freezeScreen");
                bool prepareTiltWorkbook = ConfigParsing.GetBoolYesNo(config, "prepareTiltWorkbook");
                bool computeMean = ConfigParsing.GetBoolYesNo(config, "computeMean");
                string strcomputeMeans = computeMean ? "Yes" : "No";
                #endregion

                #region License validation
                Console.WriteLine($"{headingNo++}. Validating the software license");
                string licenseCode = CleanConfig(config["LicenseCode"]);
                if (licenseCode.Length == 0)
                    throw new ConfigurationErrorsException("\nLicenseCode missing/empty.");
                LicenseValidator.ValidateLicense("TLTEXP", licenseCode);
                Console.WriteLine($"{strTab1}Validated");
                #endregion

                #region EPPlus license
                gnaT.epplusLicense();
                #endregion

                #region Workbook variables
                string strExcelPath = ConfigParsing.GetRequiredString(config, "ExcelPath");
                string strExcelFile = ConfigParsing.GetRequiredString(config, "ExcelFile");
                string strSurveyWorksheet = ConfigParsing.GetRequiredString(config, "SurveyWorksheet");
                #endregion

                #region CSV settings
                Console.WriteLine($"{headingNo++}. Variables");
                Console.WriteLine($"{strTab1}CSV settings");
                string includeHeader = CleanConfig(config["includeHeader"]);
                if (includeHeader.Length == 0) includeHeader = "Yes";

                string includeReplacementNames = CleanConfig(config["includeReplacementNames"]);
                if (includeReplacementNames.Length == 0) includeReplacementNames = "Yes";

                string includeTriggerValues = CleanConfig(config["includeTriggerValues"]);
                if (includeTriggerValues.Length == 0) includeTriggerValues = "Yes";

                string OutputFileExtension = CleanConfig(config["OutputFileExtension"]);
                if (OutputFileExtension.Length == 0) OutputFileExtension = "csv";

                string CSVseparator = CleanConfig(config["CSVseparator"]);
                if (CSVseparator.Length == 0) CSVseparator = ",";

                string CSVformat = CleanConfig(config["CSVformat"]);
                if (CSVformat.Length == 0) CSVformat = "Standard";

                string[] allowedFormats = { "Standard", "IRE"};
                if (!allowedFormats.Contains(CSVformat, StringComparer.OrdinalIgnoreCase))
                    throw new ConfigurationErrorsException($"\nCSVformat invalid. Value='{CSVformat}'. Allowed: {string.Join(", ", allowedFormats)}.");

                string sensorType = CleanConfig(config["sensorType"]);
                if (sensorType.Length == 0) CSVformat = "Fail";

                string[] allowedTypes = {"Tiltmeter", "Extensometer" };
                if (!allowedTypes.Contains(sensorType, StringComparer.OrdinalIgnoreCase))
                    throw new ConfigurationErrorsException($"\nSensorType invalid. Value='{sensorType}'. Allowed: {string.Join(", ", allowedTypes)}.");

                CsvSettings csvSettings = new CsvSettings
                {
                    IncludeHeader = includeHeader,
                    IncludeReplacementNames = includeReplacementNames,
                    IncludeTriggerValues = includeTriggerValues,
                    OutputFileExtension = OutputFileExtension,
                    CSVseparator = CSVseparator,
                    CSVformat = CSVformat,
                    SensorType = sensorType
                };

                #endregion

                #region Email settings
                Console.WriteLine($"{strTab1}Email settings");

                string strEmailLogin = CleanConfig(config["EmailLogin"]);
                string strEmailPassword = CleanConfig(config["EmailPassword"]);
                string strEmailFrom = CleanConfig(config["EmailFrom"]);
                string strEmailRecipients = CleanConfig(config["EmailRecipients"]);

                EmailCredentials emailCreds = gnaT.BuildEmailCredentials(
                    strEmailLogin,
                    strEmailPassword,
                    strEmailFrom,
                    strEmailRecipients);
                #endregion

                #region SMS settings
                Console.WriteLine($"{strTab1}SMS settings");

                List<string> smsMobile = new();
                foreach (string key in config.AllKeys.Where(k => !string.IsNullOrWhiteSpace(k) &&
                                                                k.StartsWith("RecipientPhone", StringComparison.OrdinalIgnoreCase)))
                {
                    string value = CleanConfig(config[key]);
                    if (value.Length == 0) continue;
                    smsMobile.Add(value);
                }
                #endregion

                #region General variables

                Console.WriteLine($"{strTab1}General variables");

                string strComputeMeanDeltas = CleanConfig(config["computeMean"]);
                if (strComputeMeanDeltas.Length == 0) strComputeMeanDeltas = "No";

                string strUpdateSensorList = CleanConfig(config["updateSensorList"]);
                if (strUpdateSensorList.Length == 0) strUpdateSensorList = "No";

                strSystemLogsFolder = CleanConfig(config["SystemLogsFolder"]);
                if (strSystemLogsFolder.Length == 0) strSystemLogsFolder = @"C:\__SystemLogs\";

                string strAlarmfolder = CleanConfig(config["SystemAlarmFolder"]);
                if (strAlarmfolder.Length == 0) strAlarmfolder = @"C:\__SystemAlarms\";

                Directory.CreateDirectory(strSystemLogsFolder);
                Directory.CreateDirectory(strAlarmfolder);

                string strTimeBlockType = CleanConfig(config["TimeBlockType"]);
                if (strTimeBlockType.Length == 0) strTimeBlockType = "Schedule";

                string strManualBlockStart = gnaT.NormalizeTimeStampToString(CleanConfig(config["manualBlockStart"]));
                string strManualBlockEnd = gnaT.NormalizeTimeStampToString(CleanConfig(config["manualBlockEnd"]));

                string strBlockSizeHrs = CleanConfig(config["BlockSizeHrs"]);
                if (strBlockSizeHrs.Length == 0) strBlockSizeHrs = "6";
                var cs = ConfigurationManager.ConnectionStrings["DBconnectionString"];
                if (cs == null || string.IsNullOrWhiteSpace(cs.ConnectionString))
                    throw new ConfigurationErrorsException("\nMissing connection string 'DBconnectionString'.");
                string strDBconnection = cs.ConnectionString;


                string strProjectTitle = GetRequired(config, "ProjectTitle");
                string strContractTitle = GetRequired(config, "ContractTitle");
               
                string strFTPSubdirectory = GetRequired(config, "FTPSubdirectory");

                int iFirstDataRow = GetRequiredInt(config, "FirstDataRow", 1, 1000000);
                string strFirstDataRow = iFirstDataRow.ToString(CultureInfo.InvariantCulture);

                int iFirstOutputRow = GetRequiredInt(config, "FirstOutputRow", 1, 1000000);
                string strFirstOutputRow = iFirstOutputRow.ToString(CultureInfo.InvariantCulture);



                string strExcelWorkbookFullPath = Path.Combine(strExcelPath, strExcelFile);
                if (!File.Exists(strExcelWorkbookFullPath))
                    throw new FileNotFoundException("Excel workbook not found.", strExcelWorkbookFullPath);


                if (!Directory.Exists(strFTPSubdirectory))
                    Directory.CreateDirectory(strFTPSubdirectory);

                string strMasterWorkbookFullPath = strExcelPath + strExcelFile;
                #endregion

                #region System environment check
                Console.WriteLine($"{headingNo++}. Check system environment");
                Console.WriteLine($"{strTab1}Project: " + strProjectTitle);
                Console.WriteLine($"{strTab1}Master workbook: " + strMasterWorkbookFullPath);
                //check the worksheets

                if (freezeScreen)
                {
                    Console.WriteLine($"{strTab1}Check database connection");
                    gnaDBAPI.testDBconnection(strDBconnection);

                    Console.WriteLine($"{strTab1}Check Existance of workbook & worksheet");
                    gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);

                    Console.WriteLine($"{strTab1}Update SensorID");
                    Console.WriteLine($"{strTab2}Read point names");
                    string[] strPointNames = gnaSpreadsheetAPI.readPointNames(
                        strMasterWorkbookFullPath,
                        strSurveyWorksheet,
                        iFirstDataRow.ToString(System.Globalization.CultureInfo.InvariantCulture));

                    Console.WriteLine($"{strTab2}Extract SensorID");
                    string[,] strSensorID = gnaDBAPI.getSensorIDfromDB(strDBconnection, strPointNames, strProjectTitle);

                    Console.WriteLine($"{strTab2}Update SensorID");
                    gnaSpreadsheetAPI.writeSensorID(
                        strMasterWorkbookFullPath,
                        strSurveyWorksheet,
                        strSensorID,
                        iFirstDataRow.ToString(System.Globalization.CultureInfo.InvariantCulture));


                    // update the SendorID list in the config file if needed
                    Console.WriteLine($"{strTab1}Done");
                    
                }
                else
                {
                    Console.WriteLine($"{strTab1}Skip");
                }

                #endregion

                #region Populate the RuntimeEnvironment class
                Console.WriteLine($"{headingNo++}. Populate Runtime class");
                RuntimeEnvironment runtimeEnvironment = gnaSpreadsheetAPI.populateRuntimeClass(
                    dbConnectionString: strDBconnection,
                    projectTitle: strProjectTitle,
                    excelPath: strExcelPath,
                    excelFile: strExcelFile,
                    surveyWorksheet: strSurveyWorksheet,
                    firstDataRow: iFirstDataRow,
                    firstDataCol: 1,
                    firstOutputRow: iFirstOutputRow);
                Console.WriteLine($"{strTab1}RuntimeEnvironment populated");

                #endregion

                #region Populate the Sensor list
                Console.WriteLine($"{headingNo++}. Populate sensor list");

                List<Sensor> sensorsList = gnaSpreadsheetAPI.readSensors(
                    runtimeEnvironment: runtimeEnvironment);

                Console.WriteLine($"{strTab1}Sensors loaded: {sensorsList.Count.ToString(CultureInfo.InvariantCulture)}");


                #endregion

                #region Run header log
                {
                    string runHeader =
                        $"Run start | Build={BuildInfo.BuildDateString()} | Project='{strProjectTitle}' | Contract='{strContractTitle}' | " +
                        $"Mode={(prepareTiltWorkbook ? "PrepareWorkbook" : "Export")} | TimeBlockType='{strTimeBlockType}' | " +
                        $"ManualStart='{strManualBlockStart}' | ManualEnd='{strManualBlockEnd}' | BlockSizeHrs='{strBlockSizeHrs}' | " +
                        $"ComputeMeanDeltas='{strComputeMeanDeltas}' | CSVformat='{CSVformat}' | Sep='{CSVseparator}' | Header='{includeHeader}' | " +
                        $"Workbook='{strExcelWorkbookFullPath}' | SurveyWS='{strSurveyWorksheet}' | FirstRow={iFirstDataRow} | " +
                        $"OutDir='{strFTPSubdirectory}'";
                    gnaT.updateSystemLogFile(strSystemLogsFolder, runHeader);
                }
                #endregion

                #region Timeblocks
                Console.WriteLine($"{headingNo++}. Timeblocks");
                string strTimeblockCase = strTimeBlockType.Trim().ToUpperInvariant();
                Console.WriteLine($"{strTab1}{strTimeBlockType}");

                List<Tuple<string, string>> subBlocks;
                string strManualEmailTime = "";

                // prepareTimeBlocks() returns UTC timestamps as strings (yyyy-MM-dd HH:mm:ss)
                switch (strTimeblockCase)
                {
                    case "HISTORIC":
                        subBlocks = gnaT.prepareTimeBlocks("Historic", strBlockSizeHrs, strManualBlockStart, strManualBlockEnd);
                        break;

                    case "MANUAL":
                        subBlocks = gnaT.prepareTimeBlocks("Manual", strManualBlockStart, strManualBlockEnd);
                        strManualEmailTime = BuildManualEmailTime(strManualBlockEnd);
                        break;

                    case "SCHEDULE":
                        subBlocks = gnaT.prepareTimeBlocks("Schedule", strBlockSizeHrs);
                        break;

                    default:
                        throw new ConfigurationErrorsException($"\nInvalid TimeBlockType '{strTimeBlockType}'. Must be Manual, Schedule or Historic.");
                }
                #endregion


                #region Main program
               Console.WriteLine($"{headingNo++}. Export Tilt Data: {strTimeBlockType}: {CSVformat} format");

                if (sensorsList == null || sensorsList.Count == 0)
                    throw new InvalidOperationException("Sensors list is empty.");

                Console.WriteLine($"{strTab1}Iterate over time blocks:");

                foreach (var block in subBlocks)
                {
                    string blockStartUTC = gnaT.NormalizeTimeStampToString(block.Item1);
                    string blockEndUTC = gnaT.NormalizeTimeStampToString(block.Item2);

                    // Deterministic filename: ContractTitle + formatted block end time (UTC)
                    string formattedTime = FormatUtcForFilename(blockEndUTC);
                    string expectedCsvPath = Path.Combine(
                        strFTPSubdirectory,
                        $"{SanitizeForFilename(strContractTitle)}_{formattedTime}.{OutputFileExtension}");

                    // Idempotency: skip if already exists
                    if (File.Exists(expectedCsvPath))
                    {
                        Console.WriteLine($"{strTab3}CSV exists (skip): {expectedCsvPath}");
                        gnaT.updateSystemLogFile(strSystemLogsFolder, $"Skipped (exists): {expectedCsvPath}");
                        continue;
                    }



                    List<SensorObservation> blockResults = new();

                    switch (sensorType)
                    {
                        case "Tiltmeter":
                            Console.WriteLine($"{strTab1}Sensor type: Tiltmeter");
                            Console.WriteLine($"{strTab2}Retrieving tilt data: {blockStartUTC} to {blockEndUTC}");

                            blockResults = t4dapi.GetAllSensorsAllTilts_PerTiltStart_OnePass(
                                strDBconnection,
                                sensorsList,
                                strTimeBlockType,
                                blockStartUTC,
                                blockEndUTC,
                                strcomputeMeans);

                            if (blockResults == null || blockResults.Count == 0)
                            {
                                Console.WriteLine($"{strTab3}No tilt data retrieved up to {blockEndUTC}");
                                continue;
                            }
                            else
                            {
                                Console.WriteLine($"{strTab3}Tilt data retrieved: {blockResults.Count}");
                            }

                            break;
                        case "Extensometer":
                            Console.WriteLine($"{strTab1}Sensor type: Extensometer");
                            blockResults = t4dapi.GetAllSensorsAllLengths_PerExtensometerStart_OnePass(
                                strDBconnection,
                                sensorsList,
                                strTimeBlockType,
                                blockStartUTC,
                                blockEndUTC,
                                strcomputeMeans);
                            break;
                        default:
                            throw new ConfigurationErrorsException($"\nInvalid SensorType '{sensorType}'. Allowed: Tiltmeter, Extensometer.");
                    }


                    if (strTimeblockCase == "SCHEDULE")
                    {
                        gnaSpreadsheetAPI.UpdateLastRetrievedTimeBySensor(
                            strExcelWorkbookFullPath: strMasterWorkbookFullPath,
                            strSurveyWorksheet: strSurveyWorksheet,
                            strFirstDataRow: strFirstDataRow,
                            blockResults: blockResults);

                        Console.WriteLine($"{strTab3}Timestamp updated in {strSurveyWorksheet} worksheet");
                    }
                    else
                    {
                        Console.WriteLine($"{strTab3}Not updating timestamps in workbook (TimeBlockType={strTimeBlockType})");
                    }

                    Console.WriteLine($"{strTab3}Generate CSV file:");
                    gnaT.generateSensorCSVfile(
                        csvSettings,
                        sensorsList,
                        blockResults,
                        strFTPSubdirectory,
                        strContractTitle,
                        blockEndUTC);
                    Console.WriteLine($"{strTab2}Done");
                }
                exitCode = 0;
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
                try { File.WriteAllText("fatal_crash.log", ex.ToString()); } catch { }

                try
                {
                    Console.WriteLine("Fatal crash:");
                    Console.WriteLine(ex);
                    Console.Out.Flush();
                }
                catch { }

            }
            finally
            {
                try
                {
                    if (gnaT != null)
                    {
                        gnaT.updateSystemLogFile(strSystemLogsFolder, "Run end | ExitCode=" + exitCode.ToString(CultureInfo.InvariantCulture));
                        Console.WriteLine("\nGNA_exportTiltData export completed...\n\n");
                        gnaT.freezeScreen("Yes");
                    }
                }
                catch { }
            }

        }


        #region Config helpers
        static string CleanConfig(string s) => (s ?? string.Empty).Trim().Trim('\'', '"');

        static string GetRequired(NameValueCollection cfg, string key)
        {
            string v = CleanConfig(cfg[key]);
            if (v.Length == 0)
                throw new ConfigurationErrorsException($"\nMissing/empty config key '{key}'.");
            return v;
        }

        static int GetRequiredInt(NameValueCollection cfg, string key, int minValueInclusive = int.MinValue, int maxValueInclusive = int.MaxValue)
        {
            string s = GetRequired(cfg, key);
            if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v))
                throw new ConfigurationErrorsException($"\nConfig key '{key}' is invalid (expected integer). Value='{s}'.");
            if (v < minValueInclusive || v > maxValueInclusive)
                throw new ConfigurationErrorsException($"\nConfig key '{key}' is out of range. Value={v}.");
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
                    throw new ConfigurationErrorsException($"\nMissing required appSetting: '{key}'.");

                string value = raw.Trim();

                if (value.Equals("Yes", StringComparison.OrdinalIgnoreCase)) return true;
                if (value.Equals("No", StringComparison.OrdinalIgnoreCase)) return false;

                throw new ConfigurationErrorsException(
                    $"Invalid value for appSetting '{key}\n': '{raw}'. Expected 'Yes' or 'No'.");
            }

            public static string GetRequiredString(System.Collections.Specialized.NameValueCollection appSettings, string key)
            {
                string? raw = appSettings[key];
                if (string.IsNullOrWhiteSpace(raw))
                    throw new ConfigurationErrorsException($"\nMissing or empty required appSetting: '{key}'.");
                return raw.Trim();
            }

            public static int GetRequiredInt(System.Collections.Specialized.NameValueCollection appSettings, string key)
            {
                string raw = GetRequiredString(appSettings, key);
                if (!int.TryParse(raw, out int value))
                    throw new ConfigurationErrorsException($"\nInvalid integer for appSetting '{key}': '{raw}'.");
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
                        $"\nInvalid double for appSetting '{key}': '{raw}'. Use '.' as decimal separator.");
                }
                return value;
            }
        }


        #endregion


    }
}
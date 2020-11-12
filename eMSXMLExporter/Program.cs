using EMPAPPBASELib;
using EMPCLIENTCOMMANDSLib;
using EMPTYPELIBRARYLib;
using NDesk.Options;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace eMSXMLExporter
{
    class Program
    {
        static OptionSet optionSet;

        static void Main(string[] args)
        {
            string loginName = null;
            string password = null;

            int? projectId = null;

            var targetObjectIds = new List<int>();

            string exportPath = null;

            bool showHelp = false;

            optionSet = new OptionSet()
            {
                { "u|userName=", "eMS {USERNAME} (SSO used if omitted)", v => loginName = v },
                { "p|password=", "eMS {PASSWORD} (SSO used if omitted)", v => password = v },
                { "i|projectId=", "eMS project {ID}", (int v) => projectId = v },
                { "e|exportedObjectId=", "Exported object {ID}", (int v) => targetObjectIds.Add(v) },
                { "o|outputFilePath=", "Output file {PATH}", v => exportPath = v },
                { "h|?|help", "Shows this help message", v => showHelp = v != null }
            };

            try { optionSet.Parse(args); }
            catch (Exception e)
            {
                ShowHelp(e.Message);
                return;
            }

            if (showHelp)
            {
                ShowHelp();
                return;
            }

            if (loginName == null ^ password == null)
            {
                if (loginName == null)
                {
                    ShowHelp("User name is missing");
                    return;
                }

                if (password == null)
                {
                    ShowHelp("Password is missing");
                    return;
                }
            }

            if (!projectId.HasValue)
            {
                ShowHelp("Project id is missing");
                return;
            }

            if (targetObjectIds.Count == 0)
            {
                ShowHelp("Exported object id is missing");
                return;
            }

            if (exportPath == null)
            {
                ShowHelp("Output file path is missing");
                return;
            }

            /// Login
            try
            {
                new EmpLoginCmdClass().Execute(new EmpLoginCmdParamsClass()
                {
                    LoginName = loginName,
                    Password = password,
                    ShowDialog = false
                });
            }

            catch (Exception)
            {
                HandleError("Failed to login");
                return;
            }


            /// Open project
            try
            {
                var openProjectCmdParams = new EmpOpenProjectCmdParamsClass() { ShowDialog = false };
                openProjectCmdParams.set_ProjectId(new EmpObjectKey { objectId = projectId.Value });

                new EmpOpenProjectCmdClass().Execute(openProjectCmdParams);
            }

            catch (Exception)
            {
                HandleError("Failed to open project");
                LogOut();
                return;
            }


            /// Export
            try
            {
                var appItemList = new EmpAppItemListClass();

                foreach (var targetObjectId in targetObjectIds)
                {
                    var targetAppItem = new EmpAppItemClass();
                    targetAppItem.set_Key(new EmpObjectKey() { objectId = targetObjectId });

                    appItemList.Add(targetAppItem);
                }

                new EmpExportDataCmdClass().Execute(new EmpExportCmdParamsClass()
                {
                    FileName = exportPath,
                    Selection = appItemList,
                    ShowDialog = false
                });
            }

            catch (Exception) { HandleError("Failed to export"); }


            /// Close project
            try
            {
                new EmpCloseProjectCmdClass().Execute(new EmpCloseProjectCmdParamsClass()
                {
                    RunFindCheckOut = false,
                    ShowCheckedOutNodesDialog = false,
                    ShowCloseProjectDialog = false,
                    ShowDialog = false
                });
            }

            catch (Exception) { }

            LogOut();
        }

        static void LogOut()
        {
            try { new EmpLogoutCmdClass().Execute(new EmpLogoutCmdParamsClass() { ShowDialog = false }); }
            catch (Exception) { }
        }

        static void HandleError(string message)
        {
            Console.Error.WriteLine(message);
        }

        static void ShowHelp(string errorMessage = null)
        {
            if (errorMessage != null)
            {
                HandleError(errorMessage);
                Console.WriteLine();
            }

            Console.WriteLine($"Usage: {Process.GetCurrentProcess().ProcessName} [OPTIONS]+");
            Console.WriteLine();
            Console.WriteLine("Options:");
            optionSet.WriteOptionDescriptions(Console.Out);
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Timers;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;

// Multi-version compatibility
#if ACAD2022
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ACAD2024
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#else
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
#endif

[assembly: CommandClass(typeof(AutoCADImportPlugin.Commands))]

namespace AutoCADImportPlugin
{public class AutoSyncApp : IExtensionApplication
    {
        private static Timer syncTimer;
        private static bool autoSyncEnabled = false;
        private const string SYNC_BASE_FOLDER = @"C:\RhinoAutoCADSync\";
        private static DateTime lastFileTime = DateTime.MinValue;
        private static string currentSyncFolder = "";

        public void Initialize()
        {
            Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage("\nRhino-AutoCAD Sync Plugin Loaded\n");
            
            // Initialize auto-sync timer
            syncTimer = new Timer(1000); // 1 second interval
            syncTimer.Elapsed += CheckForUpdates;
        }

        public void Terminate()
        {
            if (syncTimer != null)
            {
                syncTimer.Stop();
                syncTimer.Dispose();
            }
        }

        private static void CheckForUpdates(object sender, ElapsedEventArgs e)
        {
            if (!autoSyncEnabled || string.IsNullOrEmpty(currentSyncFolder)) return;

            try
            {
                var syncFile = Path.Combine(currentSyncFolder, "rhino_export.json");
                if (!File.Exists(syncFile)) return;

                var fileTime = File.GetLastWriteTime(syncFile);
                if (fileTime <= lastFileTime) return;

                lastFileTime = fileTime;
                
                // Import on main thread
                Application.DocumentManager.MdiActiveDocument?.SendStringToExecute("IMPORTFROMRHINO ", true, false, false);
            }
            catch (Exception ex)
            {
                // Log error silently
                System.Diagnostics.Debug.WriteLine($"Auto-sync error: {ex.Message}");
            }
        }

        public static void StartAutoSync(string syncFolder)
        {
            currentSyncFolder = syncFolder;
            autoSyncEnabled = true;
            syncTimer.Start();
        }

        public static void StopAutoSync()
        {
            autoSyncEnabled = false;
            currentSyncFolder = "";
            syncTimer.Stop();
        }

        public static string GetCurrentSyncFolder()
        {
            return currentSyncFolder;
        }
    }

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rhino;
using Rhino.Commands;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.Input;
using Rhino.Input.Custom;
using Newtonsoft.Json;
using System.Windows.Forms;

namespace RhinoExportPlugin
{
    public class RhinoToAutoCADSyncPlugIn : Rhino.PlugIns.PlugIn
    {
        public RhinoToAutoCADSyncPlugIn()
        {
            Instance = this;
        }

        public static RhinoToAutoCADSyncPlugIn Instance { get; private set; }

        protected override string LocalPlugInName => "RhinoToAutoCADSync";
    }

    public class ExportToAutoCADCommand : Command
    {
        private const string SYNC_BASE_FOLDER = @"C:\RhinoAutoCADSync\";
        private const string SETTINGS_FILE = "rhino_sync_settings.json";

        public ExportToAutoCADCommand()
        {
            Instance = this;
        }

        public static ExportToAutoCADCommand Instance { get; private set; }

        public override string EnglishName => "ExportToAutoCAD";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            try
            {
                var targetFile = GetTargetAutoCADFile(doc);
                if (string.IsNullOrEmpty(targetFile))
                    return Result.Cancel;

                if (!Directory.Exists(SYNC_BASE_FOLDER))
                {
                    Directory.CreateDirectory(SYNC_BASE_FOLDER);
                }

                var targetFolder = GetSyncFolderForTarget(targetFile);
                if (!Directory.Exists(targetFolder))
                {
                    Directory.CreateDirectory(targetFolder);
                }

                var objects = GetObjectsToExport(doc);
                if (objects.Count == 0)
                {
                    RhinoApp.WriteLine("No valid objects to export.");
                    return Result.Cancel;
                }

                var exportData = new ExportData
                {
                    Timestamp = DateTime.Now,
                    TargetFile = targetFile,
                    SourceFile = doc.Path,
                    Objects = new List<ExportObject>()
                };

                foreach (var obj in objects)
                {
                    var exportObj = ConvertRhinoObject(obj);
                    if (exportObj != null)
                        exportData.Objects.Add(exportObj);
                }

                var exportFile = Path.Combine(targetFolder, "rhino_export.json");
                var metadataFile = Path.Combine(targetFolder, "sync_metadata.json");
                
                var json = JsonConvert.SerializeObject(exportData, Formatting.Indented);
                File.WriteAllText(exportFile, json);

                var metadata = new SyncMetadata
                {
                    TargetFile = targetFile,
                    SourceFile = doc.Path,
                    LastSync = DateTime.Now,
                    ObjectCount = exportData.Objects.Count
                };
                
                var metadataJson = JsonConvert.SerializeObject(metadata, Formatting.Indented);
                File.WriteAllText(metadataFile, metadataJson);

                RememberTargetChoice(doc.Path, targetFile);

                RhinoApp.WriteLine($"Exported {exportData.Objects.Count} objects to: {Path.GetFileName(targetFile)}");
                RhinoApp.WriteLine($"Sync folder: {targetFolder}");

                return Result.Success;
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error exporting to AutoCAD: {ex.Message}");
                return Result.Failure;
            }
        }

        private string GetTargetAutoCADFile(RhinoDoc doc)
        {
            var lastTarget = GetLastTarget(doc.Path);
            
            if (!string.IsNullOrEmpty(lastTarget) && File.Exists(lastTarget))
            {
                var getOption = new GetOption();
                getOption.SetCommandPrompt($"Sync target - Last used: {Path.GetFileName(lastTarget)}");
                getOption.AddOption("UseLast");
                getOption.AddOption("ChooseNew");
                
                var result = getOption.Get();
                if (result == GetResult.Option)
                {
                    if (getOption.OptionIndex() == 1)
                    {
                        return lastTarget;
                    }
                }
            }

            return ShowFileSelectionDialog();
        }

        private string ShowFileSelectionDialog()
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "AutoCAD Files (*.dwg)|*.dwg|All Files (*.*)|*.*",
                    Title = "Select AutoCAD file to sync to",
                    InitialDirectory = GetLastUsedDirectory()
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    SaveLastUsedDirectory(Path.GetDirectoryName(openFileDialog.FileName));
                    return openFileDialog.FileName;
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error showing file dialog: {ex.Message}");
            }

            return null;
        }

        private List<RhinoObject> GetObjectsToExport(RhinoDoc doc)
        {
            var objects = new List<RhinoObject>();
            
            var getOption = new GetOption();
            getOption.SetCommandPrompt("Export all objects or selected objects?");
            getOption.AddOption("All");
            getOption.AddOption("Selected");
            
            var result = getOption.Get();
            if (result != GetResult.Option)
                return objects;

            if (getOption.OptionIndex() == 1)
            {
                foreach (var obj in doc.Objects)
                {
                    if (obj.IsValid && obj.Visible)
                        objects.Add(obj);
                }
            }
            else
            {
                var selectedObjects = doc.Objects.GetSelectedObjects(false, false);
                if (selectedObjects.Length == 0)
                {
                    RhinoApp.WriteLine("No objects selected. Please select objects first.");
                    return objects;
                }
                objects.AddRange(selectedObjects);
            }

            return objects;
        }

        private string GetSyncFolderForTarget(string targetFile)
        {
            var fileName = Path.GetFileNameWithoutExtension(targetFile);
            var safeName = MakeSafeFolderName(fileName);
            return Path.Combine(SYNC_BASE_FOLDER, safeName + "_dwg");
        }

        private string MakeSafeFolderName(string fileName)
        {
            var invalid = Path.GetInvalidFileNameChars();
            return new string(fileName.Where(c => !invalid.Contains(c)).ToArray());
        }

        private void RememberTargetChoice(string rhinoFile, string targetFile)
        {
            try
            {
                var settingsFile = Path.Combine(SYNC_BASE_FOLDER, SETTINGS_FILE);
                var settings = LoadSettings(settingsFile);
                
                if (!settings.ContainsKey("file_targets"))
                {
                    settings["file_targets"] = new Dictionary<string, object>();
                }

                var fileTargets = settings["file_targets"] as Dictionary<string, object>;
                fileTargets[rhinoFile] = new
                {
                    last_target = targetFile,
                    last_sync = DateTime.Now
                };

                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(settingsFile, json);
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error saving settings: {ex.Message}");
            }
        }

        private string GetLastTarget(string rhinoFile)
        {
            try
            {
                var settingsFile = Path.Combine(SYNC_BASE_FOLDER, SETTINGS_FILE);
                if (!File.Exists(settingsFile)) return null;

                var settings = LoadSettings(settingsFile);
                if (settings.ContainsKey("file_targets"))
                {
                    var fileTargets = settings["file_targets"] as Dictionary<string, object>;
                    if (fileTargets?.ContainsKey(rhinoFile) == true)
                    {
                        var targetInfo = JsonConvert.DeserializeObject<dynamic>(fileTargets[rhinoFile].ToString());
                        return targetInfo.last_target;
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error loading last target: {ex.Message}");
            }

            return null;
        }

        private Dictionary<string, object> LoadSettings(string settingsFile)
        {
            if (!File.Exists(settingsFile))
                return new Dictionary<string, object>();

            try
            {
                var json = File.ReadAllText(settingsFile);
                return JsonConvert.DeserializeObject<Dictionary<string, object>>(json) ?? new Dictionary<string, object>();
            }
            catch
            {
                return new Dictionary<string, object>();
            }
        }

        private string GetLastUsedDirectory()
        {
            try
            {
                var settingsFile = Path.Combine(SYNC_BASE_FOLDER, SETTINGS_FILE);
                var settings = LoadSettings(settingsFile);
                
                if (settings.ContainsKey("last_directory"))
                {
                    return settings["last_directory"].ToString();
                }
            }
            catch { }

            return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        private void SaveLastUsedDirectory(string directory)
        {
            try
            {
                var settingsFile = Path.Combine(SYNC_BASE_FOLDER, SETTINGS_FILE);
                var settings = LoadSettings(settingsFile);
                settings["last_directory"] = directory;

                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(settingsFile, json);
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error saving directory: {ex.Message}");
            }
        }

        private ExportObject ConvertRhinoObject(RhinoObject rhinoObj)
        {
            try
            {
                var exportObj = new ExportObject
                {
                    Id = rhinoObj.Id.ToString(),
                    Name = rhinoObj.Name ?? $"Object_{rhinoObj.Id}",
                    Layer = rhinoObj.Attributes.LayerIndex >= 0 ? 
                            RhinoDoc.ActiveDoc.Layers[rhinoObj.Attributes.LayerIndex].Name : "Default",
                    ObjectType = rhinoObj.ObjectType.ToString(),
                    Color = rhinoObj.Attributes.ObjectColor.ToArgb()
                };

                var geometry = rhinoObj.Geometry;
                
                if (geometry is Curve curve)
                {
                    exportObj.GeometryType = "Curve";
                    exportObj.GeometryData = SerializeCurve(curve);
                }
                else if (geometry is Surface surface)
                {
                    exportObj.GeometryType = "Surface";
                    exportObj.GeometryData = SerializeSurface(surface);
                }
                else if (geometry is Brep brep)
                {
                    if (brep.IsSolid)
                    {
                        exportObj.GeometryType = "ClosedPolysurface";
                        exportObj.GeometryData = SerializeClosedPolysurface(brep);
                    }
                    else if (brep.Faces.Count > 1)
                    {
                        exportObj.GeometryType = "OpenPolysurface";
                        exportObj.GeometryData = SerializeOpenPolysurface(brep);
                    }
                    else
                    {
                        exportObj.GeometryType = "Brep";
                        exportObj.GeometryData = SerializeBrep(brep);
                    }
                }
                else if (geometry is Mesh mesh)
                {
                    exportObj.GeometryType = "Mesh";
                    exportObj.GeometryData = SerializeMesh(mesh);
                }
                else
                {
                    exportObj.GeometryType = "Other";
                    var bbox = geometry.GetBoundingBox(true);
                    exportObj.GeometryData = new
                    {
                        BoundingBox = new
                        {
                            Min = new { X = bbox.Min.X, Y = bbox.Min.Y, Z = bbox.Min.Z },
                            Max = new { X = bbox.Max.X, Y = bbox.Max.Y, Z = bbox.Max.Z }
                        }
                    };
                }

                return exportObj;
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error converting object {rhinoObj.Id}: {ex.Message}");
                return null;
            }
        }

        private object SerializeCurve(Curve curve)
        {
            var points = new List<object>();
            
            var domain = curve.Domain;
            var sampleCount = 100;
            
            for (int i = 0; i <= sampleCount; i++)
            {
                var t = domain.ParameterAt((double)i / sampleCount);
                var point = curve.PointAt(t);
                points.Add(new { X = point.X, Y = point.Y, Z = point.Z });
            }

            return new
            {
                Points = points,
                IsClosed = curve.IsClosed,
                Degree = curve.Degree,
                Length = curve.GetLength()
            };
        }

        private object SerializeSurface(Surface surface)
        {
            var uDomain = surface.Domain(0);
            var vDomain = surface.Domain(1);
            
            return new
            {
                UDomain = new { Min = uDomain.Min, Max = uDomain.Max },
                VDomain = new { Min = vDomain.Min, Max = vDomain.Max },
                IsClosed = new { U = surface.IsClosed(0), V = surface.IsClosed(1) },
                Area = surface.GetArea()
            };
        }

        private object SerializeClosedPolysurface(Brep brep)
        {
            var faces = new List<object>();
            var edges = new List<object>();
            
            foreach (var face in brep.Faces)
            {
                var faceData = new
                {
                    Index = face.FaceIndex,
                    SurfaceData = SerializeSurface(face),
                    Area = face.GetArea(),
                    IsPlanar = face.IsPlanar(),
                    OrientationIsReversed = face.OrientationIsReversed
                };
                faces.Add(faceData);
            }

            foreach (var edge in brep.Edges)
            {
                var edgeData = new
                {
                    Index = edge.EdgeIndex,
                    StartVertex = new { X = edge.StartVertex.Location.X, Y = edge.StartVertex.Location.Y, Z = edge.StartVertex.Location.Z },
                    EndVertex = new { X = edge.EndVertex.Location.X, Y = edge.EndVertex.Location.Y, Z = edge.EndVertex.Location.Z },
                    Length = edge.GetLength()
                };
                edges.Add(edgeData);
            }

            return new
            {
                IsSolid = brep.IsSolid,
                IsManifold = brep.IsManifold,
                FaceCount = brep.Faces.Count,
                EdgeCount = brep.Edges.Count,
                VertexCount = brep.Vertices.Count,
                Volume = brep.GetVolume(),
                SurfaceArea = brep.GetArea(),
                BoundingBox = new
                {
                    Min = new { X = brep.GetBoundingBox(true).Min.X, Y = brep.GetBoundingBox(true).Min.Y, Z = brep.GetBoundingBox(true).Min.Z },
                    Max = new { X = brep.GetBoundingBox(true).Max.X, Y = brep.GetBoundingBox(true).Max.Y, Z = brep.GetBoundingBox(true).Max.Z }
                },
                Faces = faces,
                Edges = edges
            };
        }

        private object SerializeOpenPolysurface(Brep brep)
        {
            var faces = new List<object>();
            
            foreach (var face in brep.Faces)
            {
                var faceData = new
                {
                    Index = face.FaceIndex,
                    SurfaceData = SerializeSurface(face),
                    Area = face.GetArea(),
                    IsPlanar = face.IsPlanar(),
                    OrientationIsReversed = face.OrientationIsReversed
                };
                faces.Add(faceData);
            }

            return new
            {
                IsSolid = false,
                IsManifold = brep.IsManifold,
                FaceCount = brep.Faces.Count,
                EdgeCount = brep.Edges.Count,
                VertexCount = brep.Vertices.Count,
                SurfaceArea = brep.GetArea(),
                BoundingBox = new
                {
                    Min = new { X = brep.GetBoundingBox(true).Min.X, Y = brep.GetBoundingBox(true).Min.Y, Z = brep.GetBoundingBox(true).Min.Z },
                    Max = new { X = brep.GetBoundingBox(true).Max.X, Y = brep.GetBoundingBox(true).Max.Y, Z = brep.GetBoundingBox(true).Max.Z }
                },
                Faces = faces
            };
        }

        private object SerializeBrep(Brep brep)
        {
            var face = brep.Faces[0];
            
            return new
            {
                FaceCount = 1,
                EdgeCount = brep.Edges.Count,
                IsSolid = false,
                SurfaceData = SerializeSurface(face),
                Area = face.GetArea(),
                IsPlanar = face.IsPlanar(),
                BoundingBox = new
                {
                    Min = new { X = brep.GetBoundingBox(true).Min.X, Y = brep.GetBoundingBox(true).Min.Y, Z = brep.GetBoundingBox(true).Min.Z },
                    Max = new { X = brep.GetBoundingBox(true).Max.X, Y = brep.GetBoundingBox(true).Max.Y, Z = brep.GetBoundingBox(true).Max.Z }
                }
            };
        }

        private object SerializeMesh(Mesh mesh)
        {
            var vertices = new List<object>();
            var faces = new List<object>();

            foreach (var vertex in mesh.Vertices)
            {
                vertices.Add(new { X = vertex.X, Y = vertex.Y, Z = vertex.Z });
            }

            foreach (var face in mesh.Faces)
            {
                if (face.IsTriangle)
                {
                    faces.Add(new { A = face.A, B = face.B, C = face.C, Type = "Triangle" });
                }
                else
                {
                    faces.Add(new { A = face.A, B = face.B, C = face.C, D = face.D, Type = "Quad" });
                }
            }

            return new
            {
                Vertices = vertices,
                Faces = faces,
                VertexCount = mesh.Vertices.Count,
                FaceCount = mesh.Faces.Count
            };
        }
    }

    public class ExportToAutoCADLastCommand : Command
    {
        public ExportToAutoCADLastCommand()
        {
            Instance = this;
        }

        public static ExportToAutoCADLastCommand Instance { get; private set; }

        public override string EnglishName => "ExportToAutoCADLast";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            try
            {
                var exportCommand = ExportToAutoCADCommand.Instance;
                if (exportCommand == null)
                {
                    RhinoApp.WriteLine("Export command not available.");
                    return Result.Failure;
                }

                var lastTarget = GetLastTargetSilent(doc.Path);
                if (string.IsNullOrEmpty(lastTarget) || !File.Exists(lastTarget))
                {
                    RhinoApp.WriteLine("No previous sync target found. Use 'ExportToAutoCAD' to select target.");
                    return Result.Cancel;
                }

                RhinoApp.WriteLine($"Quick sync to: {Path.GetFileName(lastTarget)}");
                
                return exportCommand.RunCommand(doc, mode);
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error in quick sync: {ex.Message}");
                return Result.Failure;
            }
        }

        private string GetLastTargetSilent(string rhinoFile)
        {
            try
            {
                const string SYNC_BASE_FOLDER = @"C:\RhinoAutoCADSync\";
                const string SETTINGS_FILE = "rhino_sync_settings.json";
                
                var settingsFile = Path.Combine(SYNC_BASE_FOLDER, SETTINGS_FILE);
                if (!File.Exists(settingsFile)) return null;

                var json = File.ReadAllText(settingsFile);
                var settings = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                
                if (settings?.ContainsKey("file_targets") == true)
                {
                    var fileTargets = settings["file_targets"] as Newtonsoft.Json.Linq.JObject;
                    if (fileTargets?.ContainsKey(rhinoFile) == true)
                    {
                        var targetInfo = fileTargets[rhinoFile];
                        return targetInfo["last_target"]?.ToString();
                    }
                }
            }
            catch { }

            return null;
        }
    }

    public class ExportData
    {
        public DateTime Timestamp { get; set; }
        public string TargetFile { get; set; }
        public string SourceFile { get; set; }
        public List<ExportObject> Objects { get; set; }
    }

    public class ExportObject
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Layer { get; set; }
        public string ObjectType { get; set; }
        public string GeometryType { get; set; }
        public object GeometryData { get; set; }
        public int Color { get; set; }
    }

    public class SyncMetadata
    {
        public string TargetFile { get; set; }
        public string SourceFile { get; set; }
        public DateTime LastSync { get; set; }
        public int ObjectCount { get; set; }
    }
}

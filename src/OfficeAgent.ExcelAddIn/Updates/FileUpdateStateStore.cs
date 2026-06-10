using System;
using System.IO;
using Newtonsoft.Json;
using OfficeAgent.Core.Diagnostics;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class FileUpdateStateStore : IUpdateStateStore
    {
        private readonly string path;

        public FileUpdateStateStore(string path)
        {
            this.path = path ?? throw new ArgumentNullException(nameof(path));
        }

        public UpdateState Load()
        {
            if (!File.Exists(path))
            {
                return new UpdateState();
            }

            try
            {
                var body = File.ReadAllText(path);
                return JsonConvert.DeserializeObject<UpdateState>(body) ?? new UpdateState();
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Warn("updates", "state.load_failed", "Failed to load update state.", ex.Message);
                return new UpdateState();
            }
        }

        public void Save(UpdateState state)
        {
            string tempPath = null;
            try
            {
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var tempDirectory = string.IsNullOrEmpty(directory) ? Directory.GetCurrentDirectory() : directory;
                tempPath = Path.Combine(tempDirectory, $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
                File.WriteAllText(tempPath, JsonConvert.SerializeObject(state ?? new UpdateState(), Formatting.Indented));

                if (File.Exists(path))
                {
                    File.Replace(tempPath, path, null);
                    tempPath = null;
                }
                else
                {
                    File.Move(tempPath, path);
                    tempPath = null;
                }
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Error("updates", "state.save_failed", "Failed to save update state.", ex);
                throw;
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempPath) && File.Exists(tempPath))
                {
                    try
                    {
                        File.Delete(tempPath);
                    }
                    catch (Exception ex)
                    {
                        OfficeAgentLog.Warn("updates", "state.temp_delete_failed", "Failed to delete temporary update state file.", ex.Message);
                    }
                }
            }
        }
    }
}

using System;
using System.IO;
using Aleksej.Finance.Excel.Settings;
using Xunit;

namespace Aleksej.Finance.Excel.UnitTests
{
    // Serializes tests that mutate the global UserSettings static state / the settings file.
    [CollectionDefinition("AddIn")]
    public class AddInCollection { }
}

namespace Aleksej.Finance.Excel.UnitTests.Infra
{
    /// <summary>
    /// Saves and restores the on-disk settings.xml around tests that write settings
    /// (the ribbon handlers, UserSettings.Save/ResetToDefaults), so a developer's real
    /// settings are never clobbered.
    /// </summary>
    internal sealed class SettingsBackup : IDisposable
    {
        private readonly string _path;
        private readonly byte[]? _original;

        public SettingsBackup()
        {
            _path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Aleksej.Finance.Excel", "settings.xml");
            _original = File.Exists(_path) ? File.ReadAllBytes(_path) : null;
        }

        public void Dispose()
        {
            UserSettings.Invalidate();
            try
            {
                if (_original != null) File.WriteAllBytes(_path, _original);
                else if (File.Exists(_path)) File.Delete(_path);
            }
            catch { /* best effort restore */ }
            UserSettings.Invalidate();
        }
    }
}

using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Diagnostics;
using System.IO;
using OpenMcdf;

namespace ARVL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // Run code after the window is loaded
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Use Dispatcher to delay code until GIF starts playing
            await Dispatcher.BeginInvoke(new Action(async () =>
            {
                /* Start scanning Revit version */

                string basicFileInfoContent = ExtractBasicFileInfo(AG_Globals.RevitFilePath);

                if (string.IsNullOrEmpty(basicFileInfoContent))
                {
                    MessageBox.Show("Error E-2: Can't extract BasicFileInfo file.");
                    Application.Current.Shutdown();
                    return;
                }

                AG_Globals.FilesRevitVersion = GetRevitVersion(basicFileInfoContent);

                if (string.IsNullOrEmpty(AG_Globals.FilesRevitVersion))
                {
                    MessageBox.Show("Error E-5: Can't find Revit version.");
                    Application.Current.Shutdown();
                    return;
                }

                AG_Globals.RevitProjectType = GetRevitProjectType(basicFileInfoContent);

                if (AG_Globals.RevitProjectType == RevitPRJtype.NULL)
                {
                    MessageBox.Show("Error E-9: Can't determine project type.");
                    Application.Current.Shutdown();
                    return;
                }

                (bool launchSuccess, string launchVersion) = LaunchRevitVersion(AG_Globals.FilesRevitVersion, AG_Globals.RevitFilePath);

                AG_Globals.LaunchRevitVersion = launchVersion;

                if (!launchSuccess)
                {
                    MessageBox.Show("Error E-4: This application can only recognize Revit project versions 2019 to 2026.");
                    Application.Current.Shutdown();
                    return;
                }

                /* End scanning Revit version */

            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

            try
            {
                if (AG_Globals.Arguments.Contains("//viewer"))
                {
                    /*
                     * THIS MESSAGE DOES NOT SHOW . . .
                     */

                    TextBlock_Status.Text = $"* * * Running Viewer Mode * * *";
                    await Task.Delay(1500);
                }

                TextBlock_Status.Text = $"* * * Detected Revit {AG_Globals.FilesRevitVersion} in the file * * *";
                await Task.Delay(1500);

                // Go through different scenarios depending on the project type
                if (AG_Globals.RevitProjectType == RevitPRJtype.CENTRAL)
                {
                    TextBlock_Status.Text = "* * * This is a workshared central project * * *";
                    await Task.Delay(1500);
                    TextBlock_Status.Text = "* * * New local created in C:\\ARVL * * *";
                    await Task.Delay(1500);

                    // Calculate the number of files and total size in the ARVL folder
                    string arvlFolderPath = @"C:\ARVL";

                    if (Directory.Exists(arvlFolderPath))
                    {
                        DirectoryInfo arvlDirectory = new DirectoryInfo(arvlFolderPath);
                        FileInfo[] files = arvlDirectory.GetFiles();
                        long totalSize = 0;

                        foreach (FileInfo file in files)
                            totalSize += file.Length;

                        TextBlock_Status.Text = $"* * * ARVL folder contains {files.Length} local files ({totalSize / (1024.0 * 1024 * 1024):F2} GB) * * *";
                        await Task.Delay(2500);

                        TextBlock_Status.Text = "* * * Fun Fact: Copied Central = Local * * *";
                        await Task.Delay(2500);
                    }

                    TextBlock_Status.Text = "* * * Don't forget to empty C:\\ARVL !!! * * *";
                    await Task.Delay(1500);
                }

                else if (AG_Globals.RevitProjectType == RevitPRJtype.LOCAL || AG_Globals.RevitProjectType == RevitPRJtype.NON_WORK_SHARED)
                {
                    if (AG_Globals.RevitProjectType == RevitPRJtype.LOCAL)
                    {
                        TextBlock_Status.Text = "* * * This is a workshared local project * * *";
                        await Task.Delay(1500);

                        // Calculate the number of files and total size in the ARVL folder
                        string arvlFolderPath = @"C:\ARVL";

                        if (Directory.Exists(arvlFolderPath))
                        {
                            DirectoryInfo arvlDirectory = new DirectoryInfo(arvlFolderPath);
                            FileInfo[] files = arvlDirectory.GetFiles();
                            long totalSize = 0;

                            foreach (FileInfo file in files)
                                totalSize += file.Length;

                            TextBlock_Status.Text = $"* * * ARVL folder contains {files.Length} local files ({totalSize / (1024.0 * 1024 * 1024):F2} GB) * * *";
                            await Task.Delay(2500);
                        }

                        TextBlock_Status.Text = "* * * Don't forget to empty C:\\ARVL !!! * * *";
                        await Task.Delay(1500);
                    }

                    else if (AG_Globals.RevitProjectType == RevitPRJtype.NON_WORK_SHARED)
                    {
                        TextBlock_Status.Text = "* * * This is a non-workshared project * * *";
                        await Task.Delay(1500);
                    }

                    if (AG_Globals.FilesRevitVersion == AG_Globals.LaunchRevitVersion)
                    {
                        TextBlock_Status.Text = $"* * * Revit {AG_Globals.FilesRevitVersion} found & launched * * *";
                        await Task.Delay(1500);
                    }

                    else
                    {
                        TextBlock_Status.Text = $"* * * Revit {AG_Globals.FilesRevitVersion} not found so launched {AG_Globals.LaunchRevitVersion} * * *";
                        await Task.Delay(2500);
                    }

                    if(AG_Globals.debugMode)
                        MessageBox.Show($"{AG_Globals.RevitFilePath}\n\n{AG_Globals.CentralModelPath}");  
                }

                else if (AG_Globals.RevitProjectType == RevitPRJtype.NULL)
                {
                    MessageBox.Show("Error E-6: Can't determine project type.");
                    Application.Current.Shutdown();
                    return;
                }

                else
                {
                    MessageBox.Show("Error E-7: Flying cows alert!");
                    Application.Current.Shutdown();
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error E-12: {ex.Message}");
                Application.Current.Shutdown();
                return;
            }

            Application.Current.Shutdown();
        }

        private string ExtractBasicFileInfo(string revitFilePath)
        {
            try
            {
                using (CompoundFile cf = new CompoundFile(revitFilePath))
                {
                    CFStream basicFileInfoStream = cf.RootStorage.GetStream("BasicFileInfo");
                    byte[] data = basicFileInfoStream.GetData();
                    return System.Text.Encoding.UTF8.GetString(data).Replace("\0", string.Empty);
                }
            }
            catch
            {
                return null;
            }
        }

        // Returns 4 characters after "Format: " text in the string. Null if not found.
        private string GetRevitVersion(string basicFileInfoContent)
        {
            if (string.IsNullOrEmpty(basicFileInfoContent))
                return null;

            try
            {
                string toBeSearched = "Format: ";
                int location = basicFileInfoContent.IndexOf(toBeSearched);

                if (location != -1)
                {
                    string afterFormat = basicFileInfoContent.Substring(location + toBeSearched.Length);
                    return afterFormat.Substring(0, 4);
                }

                return null;
            }
            catch
            {
                return null;
            }
        }
        
        // Returns the rest of the line after "Worksharing: " text in the string. Null if not found.
        private RevitPRJtype GetRevitProjectType(string basicFileInfoContent)
        {
            if (string.IsNullOrEmpty(basicFileInfoContent))
                return RevitPRJtype.NULL;

            try
            {
                string toBeSearched = "Worksharing: ";
                int location = basicFileInfoContent.IndexOf(toBeSearched);

                if (location != -1)
                {
                    string extract = basicFileInfoContent.Substring(location + toBeSearched.Length);
                    int endOfLine = extract.IndexOfAny(new[] { '\n', '\r' });

                    if (endOfLine != -1)
                    {
                        extract = extract.Substring(0, endOfLine);

                        if (extract == "Central")
                        {
                            AG_Globals.CentralModelPath = GetCentralPath(basicFileInfoContent);

                            if (string.IsNullOrEmpty(AG_Globals.CentralModelPath))
                            {
                                MessageBox.Show("Error E-10: Can't find central model path.");
                                Application.Current.Shutdown();
                                return RevitPRJtype.NULL;
                            }

                            if (AG_Globals.RevitFilePath == AG_Globals.CentralModelPath)
                                return RevitPRJtype.CENTRAL;

                            else
                                return RevitPRJtype.LOCAL;
                        }

                        else if (extract == "Not enabled")
                            return RevitPRJtype.NON_WORK_SHARED;

                        else
                            return RevitPRJtype.NULL;
                    }

                }

                return RevitPRJtype.NULL;
            }

            catch
            {
                return RevitPRJtype.NULL;
            }
        }

        // Returns the rest of the line after "Central Model Path: " text in the string. Null if not found.
        private string GetCentralPath(string basicFileInfoContent)
        {
            if (string.IsNullOrEmpty(basicFileInfoContent))
                return null;

            try
            {
                string toBeSearched = "Central Model Path: ";
                int location = basicFileInfoContent.IndexOf(toBeSearched);

                if (location != -1)
                {
                    string extract = basicFileInfoContent.Substring(location + toBeSearched.Length);
                    int endOfLine = extract.IndexOfAny(new[] { '\n', '\r' });

                    if (endOfLine != -1)
                        return extract.Substring(0, endOfLine);
                }

                return null;
            }

            catch
            {
                return null;
            }
        }

        private (bool, string) LaunchRevitVersion(string revitVersion, string revitFilePath)
        {
            try
            {
                // Define the Revit executable path for the requested version
                string revitExePath = $@"C:\Program Files\Autodesk\Revit {revitVersion}\Revit.exe";

                // Construct arguments for launching Revit
                string arguments = $"\"{revitFilePath}\"";

                if (!string.IsNullOrWhiteSpace(AG_Globals.Arguments))
                    arguments += " " + AG_Globals.Arguments;

                // Handle CENTRAL project type
                if (AG_Globals.RevitProjectType == RevitPRJtype.CENTRAL)
                {
                    string localFilePath = PrepareLocalFile(revitFilePath);

                    if (localFilePath == null)
                    {
                        MessageBox.Show("Error E-11: Failed to prepare local file for CENTRAL project.");
                        return (false, null);
                    }

                    AG_Globals.RevitFilePath = localFilePath;       // Update global path to the local file
                    arguments = $"\"{localFilePath}\"";             // Update arguments for the local file

                    if (!string.IsNullOrWhiteSpace(AG_Globals.Arguments))
                        arguments += " " + AG_Globals.Arguments;
                }

                // Attempt to launch the specified Revit version
                if (File.Exists(revitExePath))
                {
                    Process.Start(revitExePath, arguments);
                    return (true, revitVersion);
                }

                // Fallback to available Revit versions (2019–2026)
                var revitVersions = new[]
                {
                    ("2026", @"C:\Program Files\Autodesk\Revit 2026\Revit.exe"),
                    ("2025", @"C:\Program Files\Autodesk\Revit 2025\Revit.exe"),
                    ("2024", @"C:\Program Files\Autodesk\Revit 2024\Revit.exe"),
                    ("2023", @"C:\Program Files\Autodesk\Revit 2023\Revit.exe"),
                    ("2022", @"C:\Program Files\Autodesk\Revit 2022\Revit.exe"),
                    ("2021", @"C:\Program Files\Autodesk\Revit 2021\Revit.exe"),
                    ("2020", @"C:\Program Files\Autodesk\Revit 2020\Revit.exe"),
                    ("2019", @"C:\Program Files\Autodesk\Revit 2019\Revit.exe")
                };

                foreach (var (version, path) in revitVersions)
                {
                    if (File.Exists(path))
                    {
                        if (int.Parse(version) < int.Parse(revitVersion))
                        {
                            MessageBox.Show($"Error E-8: This is a {revitVersion} file. The latest Revit version you have installed is {version}. You can't open this file.");
                            Application.Current.Shutdown();

                            return (false, null);
                        }

                        Process.Start(path, arguments);

                        return (true, version);
                    }
                }

                return (false, null);
            }
            catch
            {
                return (false, null);
            }
        }

        // Prepare a local file for CENTRAL projects
        private string PrepareLocalFile(string revitFilePath)
        {
            try
            {
                string localDirectory = @"C:\ARVL";
                Directory.CreateDirectory(localDirectory); // Ensures the directory exists

                string fileName = Path.GetFileName(revitFilePath);
                string localFilePath = Path.Combine(localDirectory, fileName);

                // Avoid overwriting files by appending a suffix
                int suffix = 1;

                while (File.Exists(localFilePath))
                {
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
                    string extension = Path.GetExtension(fileName);
                    localFilePath = Path.Combine(localDirectory, $"{fileNameWithoutExtension}_{suffix++}{extension}");
                }

                File.Copy(revitFilePath, localFilePath);

                return localFilePath;
            }
            catch
            {
                return null; // Return null if file preparation fails
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

    }
}



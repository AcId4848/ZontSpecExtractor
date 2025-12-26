using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Linq;

namespace ZontSpecExtractor
{
    /// <summary>
    /// Сервис для проверки и установки обновлений через GitHub Releases API
    /// </summary>
    public class GitHubUpdateService
    {
        private const string GitHubApiBaseUrl = "https://api.github.com";
        private const string GitHubReleasesBaseUrl = "https://github.com";
        private readonly string _repositoryOwner;
        private readonly string _repositoryName;
        private readonly HttpClient _httpClient;

        public GitHubUpdateService(string repositoryOwner, string repositoryName)
        {
            _repositoryOwner = repositoryOwner ?? throw new ArgumentNullException(nameof(repositoryOwner));
            _repositoryName = repositoryName ?? throw new ArgumentNullException(nameof(repositoryName));
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "ZontSpecExtractor-UpdateChecker");
            _httpClient.Timeout = TimeSpan.FromSeconds(10);
        }

        /// <summary>
        /// Получает текущую версию приложения из .csproj файла
        /// </summary>
        public static Version GetCurrentVersion()
        {
            try
            {
                // Ищем .csproj файл в директории приложения
                var appDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (string.IsNullOrEmpty(appDirectory))
                {
                    appDirectory = AppDomain.CurrentDomain.BaseDirectory;
                }

                var csprojFiles = Directory.GetFiles(appDirectory, "*.csproj", SearchOption.TopDirectoryOnly);
                if (csprojFiles.Length == 0)
                {
                    // Пробуем найти в родительской директории (для разработки)
                    var parentDir = Directory.GetParent(appDirectory)?.FullName;
                    if (!string.IsNullOrEmpty(parentDir))
                    {
                        csprojFiles = Directory.GetFiles(parentDir, "*.csproj", SearchOption.TopDirectoryOnly);
                    }
                }

                if (csprojFiles.Length > 0)
                {
                    var csprojPath = csprojFiles[0];
                    var version = GetVersionFromCsproj(csprojPath);
                    if (version != null)
                    {
                        return version;
                    }
                }

                // Fallback: пробуем получить из AssemblyInfo
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyVersion = assembly.GetName().Version;
                return assemblyVersion ?? new Version(1, 0, 0, 0);
            }
            catch (Exception ex)
            {
                LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "GetCurrentVersion", 
                    $"Failed to get version from csproj: {ex.Message}");
                // Fallback: пробуем получить из AssemblyInfo
                try
                {
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    var version = assembly.GetName().Version;
                    return version ?? new Version(1, 0, 0, 0);
                }
                catch
                {
                    return new Version(1, 0, 0, 0);
                }
            }
        }

        /// <summary>
        /// Извлекает версию из .csproj файла (поддерживает как старый формат с namespace, так и новый SDK-style)
        /// </summary>
        private static Version? GetVersionFromCsproj(string csprojPath)
        {
            try
            {
                var doc = XDocument.Load(csprojPath);
                
                // Для SDK-style проектов (без namespace) ищем напрямую
                var versionElement = doc.Descendants().FirstOrDefault(e => 
                    e.Name.LocalName == "Version");
                
                if (versionElement != null && !string.IsNullOrWhiteSpace(versionElement.Value))
                {
                    return ParseVersionFromTag(versionElement.Value.Trim());
                }

                // Если Version не найден, ищем AssemblyVersion
                var assemblyVersionElement = doc.Descendants().FirstOrDefault(e => 
                    e.Name.LocalName == "AssemblyVersion");
                
                if (assemblyVersionElement != null && !string.IsNullOrWhiteSpace(assemblyVersionElement.Value))
                {
                    return ParseVersionFromTag(assemblyVersionElement.Value.Trim());
                }

                // Для старых проектов с namespace
                XNamespace ns = "http://schemas.microsoft.com/developer/msbuild/2003";
                var versionWithNs = doc.Descendants(ns + "Version").FirstOrDefault();
                if (versionWithNs != null && !string.IsNullOrWhiteSpace(versionWithNs.Value))
                {
                    return ParseVersionFromTag(versionWithNs.Value.Trim());
                }

                var assemblyVersionWithNs = doc.Descendants(ns + "AssemblyVersion").FirstOrDefault();
                if (assemblyVersionWithNs != null && !string.IsNullOrWhiteSpace(assemblyVersionWithNs.Value))
                {
                    return ParseVersionFromTag(assemblyVersionWithNs.Value.Trim());
                }

                return null;
            }
            catch (Exception ex)
            {
                LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "GetVersionFromCsproj", 
                    $"Failed to parse csproj: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Проверяет наличие обновлений через GitHub Releases API (стандартный способ)
        /// </summary>
        public async Task<UpdateInfo?> CheckForUpdatesAsync()
        {
            try
            {
                // Используем GitHub Releases API для получения последнего релиза
                var apiUrl = $"{GitHubApiBaseUrl}/repos/{_repositoryOwner}/{_repositoryName}/releases/latest";
                LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "CheckForUpdatesAsync", 
                    $"Проверяю обновления через GitHub API: {apiUrl}");

                var response = await _httpClient.GetStringAsync(apiUrl);
                
                if (string.IsNullOrWhiteSpace(response))
                {
                    LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "CheckForUpdatesAsync", 
                        "Пустой ответ от GitHub API");
                    return null;
                }

                // Парсим JSON ответ
                var release = JsonSerializer.Deserialize<GitHubRelease>(response, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                if (release == null || string.IsNullOrWhiteSpace(release.TagName))
                {
                    LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "CheckForUpdatesAsync", 
                        "Не удалось распарсить информацию о релизе");
                    return null;
                }

                // Извлекаем версию из тега (формат: "v1.2.3" или "1.2.3")
                var latestVersion = ParseVersionFromTag(release.TagName);
                var currentVersion = GetCurrentVersion();

                LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "CheckForUpdatesAsync", 
                    $"Текущая версия: {currentVersion}, Последняя версия: {latestVersion}");

                if (latestVersion > currentVersion)
                {
                    // Ищем asset для скачивания (.exe или .zip)
                    string? downloadUrl = null;
                    if (release.Assets != null && release.Assets.Count > 0)
                    {
                        // Ищем .exe файл
                        var exeAsset = release.Assets.FirstOrDefault(a => 
                            a.Name?.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) == true);
                        
                        if (exeAsset != null)
                        {
                            downloadUrl = exeAsset.BrowserDownloadUrl;
                        }
                        else
                        {
                            // Если .exe нет, ищем .zip
                            var zipAsset = release.Assets.FirstOrDefault(a => 
                                a.Name?.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) == true);
                            if (zipAsset != null)
                            {
                                downloadUrl = zipAsset.BrowserDownloadUrl;
                            }
                        }
                    }

                    // Если не нашли asset, формируем стандартный URL
                    if (string.IsNullOrEmpty(downloadUrl))
                    {
                        downloadUrl = $"{GitHubReleasesBaseUrl}/{_repositoryOwner}/{_repositoryName}/releases/latest/download/ZontSpecExtractor.exe";
                    }

                    var releaseUrl = !string.IsNullOrEmpty(release.HtmlUrl) 
                        ? release.HtmlUrl 
                        : $"{GitHubReleasesBaseUrl}/{_repositoryOwner}/{_repositoryName}/releases/latest";

                    var releaseNotes = !string.IsNullOrWhiteSpace(release.Body) 
                        ? release.Body 
                        : $"Доступна новая версия {latestVersion}";

                    LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "CheckForUpdatesAsync", 
                        $"Найдено обновление: {latestVersion}");

                    return new UpdateInfo
                    {
                        Version = latestVersion,
                        TagName = release.TagName,
                        ReleaseNotes = releaseNotes,
                        DownloadUrl = downloadUrl,
                        ReleaseUrl = releaseUrl
                    };
                }

                return null;
            }
            catch (HttpRequestException httpEx)
            {
                LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "CheckForUpdatesAsync", 
                    $"HTTP ошибка при проверке обновлений: {httpEx.Message}");
                return null;
            }
            catch (Exception ex)
            {
                LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "CheckForUpdatesAsync", 
                    $"Ошибка при проверке обновлений: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Скачивает и устанавливает обновление
        /// </summary>
        public async Task<bool> DownloadAndInstallUpdateAsync(UpdateInfo updateInfo, string settingsFilePath)
        {
            try
            {
                if (string.IsNullOrEmpty(updateInfo.DownloadUrl))
                {
                    MessageBox.Show("Не удалось получить URL для скачивания обновления.", 
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // Временный путь для скачивания
                var tempDir = Path.Combine(Path.GetTempPath(), "ZontSpecExtractor_Update");
                Directory.CreateDirectory(tempDir);
                
                var downloadPath = Path.Combine(tempDir, "update.zip");
                var exePath = Path.Combine(tempDir, "ZontSpecExtractor.exe");

                // Скачиваем файл
                LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "DownloadAndInstallUpdateAsync", 
                    "Downloading update...");
                
                var fileBytes = await _httpClient.GetByteArrayAsync(updateInfo.DownloadUrl);
                await File.WriteAllBytesAsync(downloadPath, fileBytes);

                // Если это ZIP, распаковываем
                if (downloadPath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    ZipFile.ExtractToDirectory(downloadPath, tempDir, true);
                }
                else
                {
                    // Если это прямой exe, просто копируем
                    File.Copy(downloadPath, exePath, true);
                }

                // Находим exe в распакованных файлах
                if (!File.Exists(exePath))
                {
                    var exeFiles = Directory.GetFiles(tempDir, "*.exe", SearchOption.AllDirectories);
                    if (exeFiles.Length > 0)
                    {
                        exePath = exeFiles[0];
                    }
                    else
                    {
                        MessageBox.Show("Не найден исполняемый файл в обновлении.", 
                            "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }

                // Сохраняем настройки во временный файл
                var settingsBackup = Path.Combine(tempDir, "app_settings.json.backup");
                if (File.Exists(settingsFilePath))
                {
                    File.Copy(settingsFilePath, settingsBackup, true);
                }

                // Создаем скрипт для обновления
                var updateScript = CreateUpdateScript(
                    exePath,
                    Application.ExecutablePath,
                    settingsBackup,
                    settingsFilePath,
                    tempDir
                );

                var scriptPath = Path.Combine(tempDir, "update.bat");
                await File.WriteAllTextAsync(scriptPath, updateScript);

                // Запускаем скрипт обновления
                var processInfo = new ProcessStartInfo
                {
                    FileName = scriptPath,
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                Process.Start(processInfo);

                return true;
            }
            catch (Exception ex)
            {
                LoggingSystem.LogException("GitHubUpdateService", "DownloadAndInstallUpdateAsync", ex);
                MessageBox.Show($"Ошибка при установке обновления: {ex.Message}", 
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private string CreateUpdateScript(string newExePath, string currentExePath, 
            string settingsBackup, string settingsTarget, string tempDir)
        {
            // Скрипт будет ждать закрытия текущего приложения, затем заменит exe и восстановит настройки
            return "@echo off\r\n" +
                   "timeout /t 2 /nobreak >nul\r\n" +
                   $"taskkill /F /IM \"{Path.GetFileName(currentExePath)}\" >nul 2>&1\r\n" +
                   "timeout /t 1 /nobreak >nul\r\n" +
                   $"copy /Y \"{newExePath}\" \"{currentExePath}\" >nul\r\n" +
                   $"if exist \"{settingsBackup}\" (\r\n" +
                   $"    copy /Y \"{settingsBackup}\" \"{settingsTarget}\" >nul\r\n" +
                   ")\r\n" +
                   $"start \"\" \"{currentExePath}\"\r\n" +
                   $"rmdir /S /Q \"{tempDir}\" >nul 2>&1\r\n";
        }

        private static Version ParseVersionFromTag(string tag)
        {
            try
            {
                // Убираем префикс "v" если есть
                var versionString = tag.TrimStart('v', 'V');
                return Version.Parse(versionString);
            }
            catch
            {
                return new Version(1, 0, 0, 0);
            }
        }

        private string GetDownloadUrl(GitHubRelease release)
        {
            // Этот метод больше не используется, но оставляем для совместимости
            // Ищем asset с расширением .exe или .zip
            if (release.Assets != null && release.Assets.Count > 0)
            {
                foreach (var asset in release.Assets)
                {
                    var fileName = asset.Name?.ToLowerInvariant() ?? "";
                    if (fileName.EndsWith(".exe") || fileName.EndsWith(".zip"))
                    {
                        return asset.BrowserDownloadUrl ?? "";
                    }
                }
            }

            // Если не нашли asset, возвращаем пустую строку
            return "";
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }

    /// <summary>
    /// Информация об обновлении
    /// </summary>
    public class UpdateInfo
    {
        public Version Version { get; set; }
        public string TagName { get; set; } = "";
        public string ReleaseNotes { get; set; } = "";
        public string DownloadUrl { get; set; } = "";
        public string ReleaseUrl { get; set; } = "";
    }

    /// <summary>
    /// Модель GitHub Release
    /// </summary>
    internal class GitHubRelease
    {
        public string TagName { get; set; } = "";
        public string Body { get; set; } = "";
        public string HtmlUrl { get; set; } = "";
        public List<GitHubAsset>? Assets { get; set; }
    }

    /// <summary>
    /// Модель GitHub Asset
    /// </summary>
    internal class GitHubAsset
    {
        public string Name { get; set; } = "";
        public string BrowserDownloadUrl { get; set; } = "";
    }
}


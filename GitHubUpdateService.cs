using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
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
        private readonly object _lockObject = new object();

        private bool _disposed = false;

        public GitHubUpdateService(string repositoryOwner, string repositoryName)
        {
            _repositoryOwner = repositoryOwner ?? throw new ArgumentNullException(nameof(repositoryOwner));
            _repositoryName = repositoryName ?? throw new ArgumentNullException(nameof(repositoryName));
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "ZontSpecExtractor-UpdateChecker");
            _httpClient.Timeout = TimeSpan.FromMinutes(5); // Увеличиваем timeout для больших файлов
        }

        /// <summary>
        /// Получает текущую версию приложения из Assembly (свойств)
        /// </summary>
        public static Version GetCurrentVersion()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyVersion = assembly.GetName().Version;
                return assemblyVersion ?? new Version(1, 0, 0, 0);
            }
            catch (Exception ex)
            {
                LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "GetCurrentVersion", 
                    $"Failed to get version from Assembly: {ex.Message}");
                return new Version(1, 0, 0, 0);
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
        /// Получает версию из .csproj файла в релизе через raw ссылку
        /// </summary>
        private async Task<Version?> GetVersionFromReleaseCsprojAsync(string tagName)
        {
            try
            {
                // Пробуем разные пути к .csproj файлу в релизе
                var possiblePaths = new[]
                {
                    $"https://raw.githubusercontent.com/{_repositoryOwner}/{_repositoryName}/{tagName}/ZontSpecExtractor.csproj",
                    $"https://raw.githubusercontent.com/{_repositoryOwner}/{_repositoryName}/refs/heads/{tagName}/ZontSpecExtractor.csproj",
                    $"https://raw.githubusercontent.com/{_repositoryOwner}/{_repositoryName}/master/ZontSpecExtractor.csproj",
                    $"https://raw.githubusercontent.com/{_repositoryOwner}/{_repositoryName}/main/ZontSpecExtractor.csproj"
                };

                HttpClient? clientToUse = null;
                lock (_lockObject)
                {
                    if (_disposed || _httpClient == null)
                    {
                        return null;
                    }
                    clientToUse = _httpClient;
                }

                foreach (var path in possiblePaths)
                {
                    try
                    {
                        var response = await clientToUse.GetStringAsync(path);
                        if (!string.IsNullOrWhiteSpace(response))
                        {
                            var version = GetVersionFromCsprojContent(response);
                            if (version != null)
                            {
                                LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "GetVersionFromReleaseCsprojAsync", 
                                    $"Версия найдена в .csproj по пути: {path}");
                                return version;
                            }
                        }
                    }
                    catch
                    {
                        // Пробуем следующий путь
                        continue;
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Извлекает версию из содержимого .csproj файла
        /// </summary>
        private static Version? GetVersionFromCsprojContent(string csprojContent)
        {
            try
            {
                var doc = XDocument.Parse(csprojContent);
                
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

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Проверяет наличие обновлений через GitHub Releases API
        /// </summary>
        public async Task<UpdateInfo?> CheckForUpdatesAsync()
        {
            try
            {
                // Получаем текущую версию из Assembly
                var currentVersion = GetCurrentVersion();
                LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "CheckForUpdatesAsync", 
                    $"Текущая версия из Assembly: {currentVersion}");

                // Используем GitHub Releases API для получения последнего релиза
                var apiUrl = $"{GitHubApiBaseUrl}/repos/{_repositoryOwner}/{_repositoryName}/releases/latest";
                LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "CheckForUpdatesAsync", 
                    $"Проверяю обновления через GitHub API: {apiUrl}");

                lock (_lockObject)
                {
                    if (_disposed || _httpClient == null)
                    {
                        LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "CheckForUpdatesAsync", 
                            "Сервис обновлений был disposed");
                        return null;
                    }
                }

                var httpResponse = await _httpClient.GetAsync(apiUrl);
                
                if (httpResponse.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "CheckForUpdatesAsync", 
                        "Релизы не найдены (404).");
                    return null;
                }

                if (!httpResponse.IsSuccessStatusCode)
                {
                    LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "CheckForUpdatesAsync", 
                        $"GitHub API вернул ошибку: {httpResponse.StatusCode}");
                    return null;
                }

                var response = await httpResponse.Content.ReadAsStringAsync();
                
                if (string.IsNullOrWhiteSpace(response))
                {
                    return null;
                }

                // Парсим JSON ответ
                GitHubRelease? release = null;
                try
                {
                    release = JsonSerializer.Deserialize<GitHubRelease>(response, new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    });
                }
                catch (JsonException jsonEx)
                {
                    LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "CheckForUpdatesAsync", 
                        $"Ошибка десериализации JSON: {jsonEx.Message}");
                    return null;
                }

                if (release == null || string.IsNullOrWhiteSpace(release.TagName))
                {
                    return null;
                }

                // Пытаемся получить версию из .csproj файла в релизе через raw ссылку
                var latestVersion = await GetVersionFromReleaseCsprojAsync(release.TagName);
                
                // Если не удалось получить из .csproj, пробуем из тега
                if (latestVersion == null)
                {
                    latestVersion = ParseVersionFromTag(release.TagName);
                }

                if (latestVersion == null)
                {
                    LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "CheckForUpdatesAsync", 
                        "Не удалось определить версию из релиза");
                    return null;
                }

                LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "CheckForUpdatesAsync", 
                    $"Текущая версия: {currentVersion}, Версия в релизе: {latestVersion}");

                var releaseUrl = !string.IsNullOrEmpty(release.HtmlUrl) 
                    ? release.HtmlUrl 
                    : $"{GitHubReleasesBaseUrl}/{_repositoryOwner}/{_repositoryName}/releases/latest";

                var releaseNotes = !string.IsNullOrWhiteSpace(release.Body) 
                    ? release.Body 
                    : $"Доступна новая версия {latestVersion}";

                // Если версия в релизе больше текущей - есть обновление
                if (latestVersion > currentVersion)
                {
                    // Ищем asset для скачивания - ОБЯЗАТЕЛЬНО ZIP (нужны все файлы)
                    string? downloadUrl = null;
                    if (release.Assets != null && release.Assets.Count > 0)
                    {
                        // Ищем .zip архив со всеми файлами (исключаем исходники)
                        var zipAsset = release.Assets.FirstOrDefault(a => 
                            a.Name?.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) == true &&
                            !a.Name.Contains("Source code", StringComparison.OrdinalIgnoreCase));
                        if (zipAsset != null)
                        {
                            downloadUrl = zipAsset.BrowserDownloadUrl;
                            LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "CheckForUpdatesAsync", 
                                $"Найден ZIP архив для скачивания: {zipAsset.Name}");
                        }
                    }

                    // Если не нашли ZIP в assets, пробуем стандартные варианты URL
                    if (string.IsNullOrEmpty(downloadUrl))
                    {
                        var possibleUrls = new[]
                        {
                            $"{GitHubReleasesBaseUrl}/{_repositoryOwner}/{_repositoryName}/releases/download/{release.TagName}/ZontSpecExtractor.zip",
                            $"{GitHubReleasesBaseUrl}/{_repositoryOwner}/{_repositoryName}/releases/latest/download/ZontSpecExtractor.zip"
                        };
                        
                        // Проверяем первый URL (с конкретным тегом)
                        downloadUrl = possibleUrls[0];
                        LoggingSystem.Log(LogLevel.WARNING, "GitHubUpdateService", "CheckForUpdatesAsync", 
                            $"ZIP архив не найден в assets, используем стандартный URL: {downloadUrl}");
                    }

                    if (string.IsNullOrEmpty(downloadUrl))
                    {
                        LoggingSystem.Log(LogLevel.ERROR, "GitHubUpdateService", "CheckForUpdatesAsync", 
                            "Не найден ZIP архив в релизе. Для обновления необходим ZIP архив со всеми файлами.");
                        return new UpdateInfo
                        {
                            Version = latestVersion,
                            TagName = release.TagName,
                            ReleaseNotes = releaseNotes,
                            DownloadUrl = "",
                            ReleaseUrl = releaseUrl,
                            HasUpdate = false // Помечаем как нет обновления, если нет ZIP
                        };
                    }

                    LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "CheckForUpdatesAsync", 
                        $"Найдено обновление: {latestVersion}");

                    return new UpdateInfo
                    {
                        Version = latestVersion,
                        TagName = release.TagName,
                        ReleaseNotes = releaseNotes,
                        DownloadUrl = downloadUrl,
                        ReleaseUrl = releaseUrl,
                        HasUpdate = true
                    };
                }
                else
                {
                    // Версия в релизе меньше или равна текущей - у пользователя актуальная версия
                    LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "CheckForUpdatesAsync", 
                        $"У пользователя актуальная версия. Текущая: {currentVersion}, В релизе: {latestVersion}");

                    return new UpdateInfo
                    {
                        Version = latestVersion,
                        TagName = release.TagName,
                        ReleaseNotes = releaseNotes,
                        DownloadUrl = "",
                        ReleaseUrl = releaseUrl,
                        HasUpdate = false
                    };
                }
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
                
                // ОБЯЗАТЕЛЬНО должен быть ZIP архив со всеми файлами
                var downloadPath = Path.Combine(tempDir, "update.zip");

                // Скачиваем файл
                LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "DownloadAndInstallUpdateAsync", 
                    $"Downloading update from: {updateInfo.DownloadUrl}");
                
                HttpClient? clientToUse = null;
                lock (_lockObject)
                {
                    if (_disposed || _httpClient == null)
                    {
                        throw new ObjectDisposedException(nameof(GitHubUpdateService), "Сервис обновлений был disposed");
                    }
                    clientToUse = _httpClient;
                }
                
                // Используем локальную ссылку вне lock, чтобы избежать проблем с disposed
                byte[] fileBytes;
                try
                {
                    fileBytes = await clientToUse.GetByteArrayAsync(updateInfo.DownloadUrl);
                }
                catch (HttpRequestException httpEx) when (httpEx.Message.Contains("404"))
                {
                    LoggingSystem.Log(LogLevel.ERROR, "GitHubUpdateService", "DownloadAndInstallUpdateAsync", 
                        $"ZIP архив не найден по URL: {updateInfo.DownloadUrl}. Убедитесь, что в релизе есть ZIP архив со всеми файлами.");
                    MessageBox.Show($"ZIP архив не найден в релизе.\n\nURL: {updateInfo.DownloadUrl}\n\nУбедитесь, что в релизе загружен ZIP архив со всеми файлами приложения.", 
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                
                await File.WriteAllBytesAsync(downloadPath, fileBytes);

                LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "DownloadAndInstallUpdateAsync", 
                    $"Скачано {fileBytes.Length} байт");

                // Проверяем, что это действительно ZIP (по магическим байтам)
                bool isActuallyZip = false;
                if (fileBytes.Length >= 4)
                {
                    // ZIP файлы начинаются с "PK" (0x50 0x4B)
                    isActuallyZip = fileBytes[0] == 0x50 && fileBytes[1] == 0x4B;
                }

                if (!isActuallyZip)
                {
                    throw new Exception("Скачанный файл не является ZIP архивом. Для обновления необходим ZIP архив со всеми файлами.");
                }

                // Распаковываем ZIP архив
                try
                {
                    ZipFile.ExtractToDirectory(downloadPath, tempDir, true);
                    LoggingSystem.Log(LogLevel.INFO, "GitHubUpdateService", "DownloadAndInstallUpdateAsync", 
                        "ZIP файл успешно распакован");
                }
                catch (Exception zipEx)
                {
                    LoggingSystem.Log(LogLevel.ERROR, "GitHubUpdateService", "DownloadAndInstallUpdateAsync", 
                        $"Ошибка при распаковке ZIP: {zipEx.Message}");
                    throw new Exception($"Не удалось распаковать архив обновления. Возможно, файл поврежден.", zipEx);
                }

                // Находим exe в распакованных файлах (для скрипта)
                var exeFiles = Directory.GetFiles(tempDir, "*.exe", SearchOption.AllDirectories);
                string? exePath = null;
                if (exeFiles.Length > 0)
                {
                    // Ищем основной exe файл приложения
                    exePath = exeFiles.FirstOrDefault(f => 
                        Path.GetFileName(f).Equals("ZontSpecExtractor.exe", StringComparison.OrdinalIgnoreCase) ||
                        Path.GetFileName(f).Equals("apphost.exe", StringComparison.OrdinalIgnoreCase));
                    
                    if (exePath == null)
                    {
                        exePath = exeFiles[0]; // Берем первый найденный
                    }
                }
                else
                {
                    MessageBox.Show("Не найден исполняемый файл в обновлении.", 
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
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
            var currentExeDir = Path.GetDirectoryName(currentExePath) ?? "";
            var settingsFileName = Path.GetFileName(settingsTarget);
            
            // Скрипт будет ждать закрытия текущего приложения, затем заменит ВСЕ файлы из распакованного релиза в папку приложения
            return "@echo off\r\n" +
                   "timeout /t 2 /nobreak >nul\r\n" +
                   $"taskkill /F /IM \"{Path.GetFileName(currentExePath)}\" >nul 2>&1\r\n" +
                   "timeout /t 1 /nobreak >nul\r\n" +
                   // Удаляем appsettings.json из папки обновления перед копированием (если есть)
                   $"if exist \"{Path.Combine(tempDir, settingsFileName)}\" (\r\n" +
                   $"    del /F /Q \"{Path.Combine(tempDir, settingsFileName)}\" >nul 2>&1\r\n" +
                   ")\r\n" +
                   // Копируем ВСЕ файлы из распакованного релиза в папку приложения
                   // xcopy /Y /E /I копирует все файлы и папки, /Y перезаписывает без подтверждения
                   $"xcopy /Y /E /I \"{tempDir}\" \"{currentExeDir}\" >nul 2>&1\r\n" +
                   // Восстанавливаем настройки из бэкапа (перезаписываем appsettings.json если он был скопирован)
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
            lock (_lockObject)
            {
                if (!_disposed)
                {
                    _httpClient?.Dispose();
                    _disposed = true;
                }
            }
        }
    }

    /// <summary>
    /// Информация об обновлении
    /// </summary>
    public class UpdateInfo
    {
        public Version Version { get; set; } = new Version(1, 0, 0, 0);
        public string TagName { get; set; } = "";
        public string ReleaseNotes { get; set; } = "";
        public string DownloadUrl { get; set; } = "";
        public string ReleaseUrl { get; set; } = "";
        public bool HasUpdate { get; set; } = false; // true если есть обновление (версия в релизе больше текущей)
    }

    /// <summary>
    /// Модель GitHub Release
    /// </summary>
    internal class GitHubRelease
    {
        [JsonPropertyName("tag_name")]
        public string TagName { get; set; } = "";
        
        [JsonPropertyName("body")]
        public string Body { get; set; } = "";
        
        [JsonPropertyName("html_url")]
        public string HtmlUrl { get; set; } = "";
        [JsonPropertyName("assets")]
        public List<GitHubAsset>? Assets { get; set; }
    }

    /// <summary>
    /// Модель GitHub Asset
    /// </summary>
    internal class GitHubAsset
    {
        [JsonPropertyName("name")]
        public string Name { get; set; } = "";
        
        [JsonPropertyName("browser_download_url")]
        public string BrowserDownloadUrl { get; set; } = "";
    }
}


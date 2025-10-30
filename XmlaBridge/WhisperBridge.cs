using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Reflection;
using Whisper.net;
using Whisper.net.Ggml;
using Whisper.net.Wave;
using Xabe.FFmpeg;

public class WhisperBridge
{
    private static WhisperFactory whisperFactory;
    private static WhisperProcessor processor;
    private static readonly object lockObject = new object();
    private static bool ffmpegInitialized = false;

    // Windows API for DLL loading
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    static extern bool SetDllDirectory(string lpPathName);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern IntPtr AddDllDirectory(string NewDirectory);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern IntPtr LoadLibrary(string lpFileName);

    static WhisperBridge()
    {
        // Initialize assembly resolver to handle version conflicts
        AssemblyResolver.Initialize();

        // Add native DLL directories to search path
        try
        {
            Console.WriteLine("[Whisper.NET] Setting up native DLL search paths...");

            string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            Console.WriteLine($"[Whisper.NET] Assembly directory: {assemblyDir}");

            // First, explicitly load the Whisper native DLLs from the assembly directory
            string[] whisperDlls = new[] { "ggml-whisper.dll", "whisper.dll" };
            foreach (var dllName in whisperDlls)
            {
                string dllPath = Path.Combine(assemblyDir, dllName);
                if (File.Exists(dllPath))
                {
                    Console.WriteLine($"[Whisper.NET] Explicitly loading: {dllPath}");
                    IntPtr handle = LoadLibrary(dllPath);
                    if (handle != IntPtr.Zero)
                    {
                        Console.WriteLine($"[Whisper.NET] ✓ Successfully loaded {dllName}");
                    }
                    else
                    {
                        int errorCode = Marshal.GetLastWin32Error();
                        Console.WriteLine($"[Whisper.NET] ✗ Failed to load {dllName}, error code: {errorCode}");
                    }
                }
                else
                {
                    Console.WriteLine($"[Whisper.NET] ✗ DLL not found: {dllPath}");
                }
            }

            // Add various possible locations for native DLLs to search path
            string[] possiblePaths = new[]
            {
                assemblyDir,  // Primary location where we copied the DLLs
                Path.Combine(assemblyDir, "runtimes", "win-x64", "native"),
                Path.Combine(assemblyDir, "runtimes", "win-x64"),
                Path.Combine(assemblyDir, "runtimes", "win-x86", "native"),
                Path.Combine(assemblyDir, "runtimes", "win", "native")
            };

            foreach (var nativePath in possiblePaths)
            {
                if (Directory.Exists(nativePath))
                {
                    Console.WriteLine($"[Whisper.NET] Adding to DLL search path: {nativePath}");

                    // Try to add directory to DLL search path
                    try
                    {
                        IntPtr result = AddDllDirectory(nativePath);
                        if (result != IntPtr.Zero)
                        {
                            Console.WriteLine($"[Whisper.NET] Successfully added DLL directory: {nativePath}");
                        }
                        else
                        {
                            Console.WriteLine($"[Whisper.NET] AddDllDirectory failed for: {nativePath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Whisper.NET] Failed to add DLL directory {nativePath}: {ex.Message}");
                    }

                    // Also add to PATH environment variable as fallback
                    try
                    {
                        string currentPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                        if (!currentPath.Contains(nativePath))
                        {
                            Environment.SetEnvironmentVariable("PATH", nativePath + ";" + currentPath);
                            Console.WriteLine($"[Whisper.NET] Added to PATH: {nativePath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Whisper.NET] Failed to add to PATH {nativePath}: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"[Whisper.NET] Directory does not exist: {nativePath}");
                }
            }

            Console.WriteLine("[Whisper.NET] Native DLL search paths configured");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Whisper.NET] Error setting up native DLL paths: {ex.Message}");
        }
    }

    private static async Task EnsureFFmpegInitialized()
    {
        if (ffmpegInitialized)
        {
            return;
        }

        try
        {
            Console.WriteLine("[Whisper.NET] Initializing FFmpeg...");

            // Enable TLS 1.2 for HTTPS downloads
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            // Set FFmpeg path to a writable location
            string ffmpegPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Whisper.NET",
                "ffmpeg"
            );

            if (!Directory.Exists(ffmpegPath))
            {
                Directory.CreateDirectory(ffmpegPath);
            }

            FFmpeg.SetExecutablesPath(ffmpegPath);

            // Download FFmpeg binaries if they don't exist
            if (!File.Exists(Path.Combine(ffmpegPath, "ffmpeg.exe")))
            {
                Console.WriteLine("[Whisper.NET] Downloading FFmpeg binaries (~100MB, one-time download)...");
                await Xabe.FFmpeg.Downloader.FFmpegDownloader.GetLatestVersion(Xabe.FFmpeg.Downloader.FFmpegVersion.Official, ffmpegPath);
                Console.WriteLine("[Whisper.NET] FFmpeg download complete");
            }

            ffmpegInitialized = true;
            Console.WriteLine("[Whisper.NET] FFmpeg initialized");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Whisper.NET] FFmpeg initialization failed: {ex.Message}");
            throw;
        }
    }

    private static void DownloadModelWithRetry(string modelPath)
    {
        const int maxRetries = 3;
        const string modelUrl = "https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-base.bin";

        // Enable TLS 1.2 (required for HuggingFace)
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
        Console.WriteLine($"[Whisper.NET] TLS protocol set to: {ServicePointManager.SecurityProtocol}");

        for (int attempt = 1; attempt <= maxRetries; attempt++)
        {
            try
            {
                Console.WriteLine($"[Whisper.NET] Download attempt {attempt}/{maxRetries}...");

                using (var httpClient = new HttpClient())
                {
                    // Set timeout to 10 minutes
                    httpClient.Timeout = TimeSpan.FromMinutes(10);

                    // Add user agent
                    httpClient.DefaultRequestHeaders.Add("User-Agent", "WhisperNET-PowerBI-Chat/1.0");

                    // Download with progress
                    var downloadTask = httpClient.GetAsync(modelUrl, HttpCompletionOption.ResponseHeadersRead);
                    downloadTask.Wait();

                    using (var response = downloadTask.Result)
                    {
                        response.EnsureSuccessStatusCode();

                        var totalBytes = response.Content.Headers.ContentLength ?? -1;
                        Console.WriteLine($"[Whisper.NET] Total size: {totalBytes / 1024 / 1024}MB");

                        using (var contentStream = response.Content.ReadAsStreamAsync().Result)
                        using (var fileStream = new FileStream(modelPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                        {
                            var buffer = new byte[8192];
                            long totalRead = 0;
                            int bytesRead;
                            int lastProgress = 0;

                            while ((bytesRead = contentStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                fileStream.Write(buffer, 0, bytesRead);
                                totalRead += bytesRead;

                                if (totalBytes > 0)
                                {
                                    int progress = (int)((totalRead * 100) / totalBytes);
                                    if (progress >= lastProgress + 10)
                                    {
                                        Console.WriteLine($"[Whisper.NET] Downloaded: {progress}%");
                                        lastProgress = progress;
                                    }
                                }
                            }

                            fileStream.Flush();
                        }
                    }
                }

                Console.WriteLine($"[Whisper.NET] Download completed on attempt {attempt}");
                return; // Success!
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Whisper.NET] Download attempt {attempt} failed");
                Console.WriteLine($"[Whisper.NET] Exception type: {ex.GetType().FullName}");
                Console.WriteLine($"[Whisper.NET] Exception message: {ex.Message}");

                // Unwrap AggregateException to get the real error
                if (ex is AggregateException aggEx)
                {
                    Console.WriteLine($"[Whisper.NET] AggregateException with {aggEx.InnerExceptions.Count} inner exceptions:");
                    foreach (var innerEx in aggEx.InnerExceptions)
                    {
                        Console.WriteLine($"[Whisper.NET]   - {innerEx.GetType().FullName}: {innerEx.Message}");
                        if (innerEx.InnerException != null)
                        {
                            Console.WriteLine($"[Whisper.NET]     Inner: {innerEx.InnerException.GetType().FullName}: {innerEx.InnerException.Message}");
                        }
                    }
                }
                else if (ex.InnerException != null)
                {
                    Console.WriteLine($"[Whisper.NET] Inner exception: {ex.InnerException.GetType().FullName}");
                    Console.WriteLine($"[Whisper.NET] Inner message: {ex.InnerException.Message}");
                }

                // Clean up partial file
                if (File.Exists(modelPath))
                {
                    try
                    {
                        File.Delete(modelPath);
                    }
                    catch { }
                }

                if (attempt == maxRetries)
                {
                    var realError = ex;
                    if (ex is AggregateException agg && agg.InnerExceptions.Count > 0)
                    {
                        realError = agg.InnerExceptions[0];
                    }

                    throw new Exception(
                        $"Failed to download Whisper model after {maxRetries} attempts. " +
                        $"Error: {realError.Message}. " +
                        $"Please check your internet connection and firewall settings.",
                        ex
                    );
                }

                // Wait before retry (exponential backoff)
                var waitTime = TimeSpan.FromSeconds(Math.Pow(2, attempt));
                Console.WriteLine($"[Whisper.NET] Waiting {waitTime.TotalSeconds} seconds before retry...");
                System.Threading.Thread.Sleep(waitTime);
            }
        }
    }

    public async Task<object> Invoke(dynamic input)
    {
        try
        {
            Console.WriteLine("[Whisper.NET] === Invoke method called ===");

            // Get the audio data as base64 string
            string audioBase64 = (string)input.audioData;

            Console.WriteLine("[Whisper.NET] Starting transcription...");
            Console.WriteLine($"[Whisper.NET] Audio data length: {audioBase64?.Length ?? 0} chars");

            if (string.IsNullOrEmpty(audioBase64))
            {
                throw new ArgumentException("No audio data provided");
            }

            // Convert base64 to byte array
            byte[] audioBytes = Convert.FromBase64String(audioBase64);
            Console.WriteLine($"[Whisper.NET] Decoded audio bytes: {audioBytes.Length}");

            // Initialize Whisper model if not already loaded
            EnsureModelLoaded();

            // Initialize FFmpeg
            await EnsureFFmpegInitialized();

            // Save audio to temporary file
            string tempAudioPath = Path.Combine(Path.GetTempPath(), $"whisper_audio_{Guid.NewGuid()}.webm");
            string tempWavPath = Path.Combine(Path.GetTempPath(), $"whisper_wav_{Guid.NewGuid()}.wav");

            try
            {
                // Save the original audio
                File.WriteAllBytes(tempAudioPath, audioBytes);
                Console.WriteLine($"[Whisper.NET] Saved audio to: {tempAudioPath}");

                // Convert to WAV format that Whisper expects (16kHz, mono, 16-bit PCM)
                Console.WriteLine($"[Whisper.NET] Converting to WAV format...");
                var mediaInfo = await FFmpeg.GetMediaInfo(tempAudioPath);
                var audioStream = mediaInfo.AudioStreams.FirstOrDefault();

                if (audioStream == null)
                {
                    throw new Exception("No audio stream found in the recording");
                }

                audioStream
                    .SetSampleRate(16000)  // Whisper requires 16kHz
                    .SetChannels(1);        // Mono

                var conversion = FFmpeg.Conversions.New()
                    .AddStream(audioStream)
                    .SetOutput(tempWavPath);

                await conversion.Start();
                Console.WriteLine($"[Whisper.NET] Conversion complete: {tempWavPath}");

                string transcription = "";

                using (FileStream fileStream = File.OpenRead(tempWavPath))
                {
                    Console.WriteLine($"[Whisper.NET] Processing audio stream, size: {fileStream.Length} bytes");

                    // Process asynchronously and wait for result
                    var segments = new List<string>();

                    try
                    {
                        await foreach (var segment in processor.ProcessAsync(fileStream))
                        {
                            if (!string.IsNullOrWhiteSpace(segment.Text))
                            {
                                segments.Add(segment.Text.Trim());
                                Console.WriteLine($"[Whisper.NET] Segment: {segment.Text}");
                            }
                        }
                    }
                    catch (Exception procEx)
                    {
                        Console.WriteLine($"[Whisper.NET] Processing error: {procEx.Message}");
                        Console.WriteLine($"[Whisper.NET] Inner exception: {procEx.InnerException?.Message}");
                        throw new Exception($"Audio processing failed: {procEx.Message}. The audio format may not be compatible. Please try speaking again.", procEx);
                    }

                    transcription = string.Join(" ", segments).Trim();
                }

                if (string.IsNullOrEmpty(transcription))
                {
                    throw new Exception("No speech detected in the audio. Please try speaking more clearly or check your microphone.");
                }

                Console.WriteLine($"[Whisper.NET] Transcription complete: {transcription}");

                return new
                {
                    success = true,
                    text = transcription
                };
            }
            finally
            {
                // Clean up temporary files
                try
                {
                    if (File.Exists(tempAudioPath))
                    {
                        File.Delete(tempAudioPath);
                        Console.WriteLine($"[Whisper.NET] Cleaned up temp audio file");
                    }
                    if (File.Exists(tempWavPath))
                    {
                        File.Delete(tempWavPath);
                        Console.WriteLine($"[Whisper.NET] Cleaned up temp wav file");
                    }
                }
                catch (Exception cleanupEx)
                {
                    Console.WriteLine($"[Whisper.NET] Failed to delete temp files: {cleanupEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            var errorMessage = ex.Message;
            var innerMessage = ex.InnerException?.Message;
            var innerStack = ex.InnerException?.StackTrace;

            Console.WriteLine($"[Whisper.NET] === EXCEPTION CAUGHT ===");
            Console.WriteLine($"[Whisper.NET] Error Type: {ex.GetType().FullName}");
            Console.WriteLine($"[Whisper.NET] Error: {errorMessage}");

            if (ex.InnerException != null)
            {
                Console.WriteLine($"[Whisper.NET] Inner Exception Type: {ex.InnerException.GetType().FullName}");
                Console.WriteLine($"[Whisper.NET] Inner Error: {innerMessage}");
                Console.WriteLine($"[Whisper.NET] Inner Stack: {innerStack}");
            }

            Console.WriteLine($"[Whisper.NET] Stack: {ex.StackTrace}");

            // Check for aggregate exceptions
            if (ex is AggregateException aggEx)
            {
                Console.WriteLine($"[Whisper.NET] AggregateException with {aggEx.InnerExceptions.Count} inner exceptions:");
                foreach (var innerEx in aggEx.InnerExceptions)
                {
                    Console.WriteLine($"[Whisper.NET]   - {innerEx.GetType().Name}: {innerEx.Message}");
                }
            }

            // Provide user-friendly error message
            var displayError = errorMessage;
            if (errorMessage.Contains("No such file") || errorMessage.Contains("not found"))
            {
                displayError = "Whisper model not found. It will be downloaded on first use (this may take a few minutes).";
            }
            else if (errorMessage.Contains("format"))
            {
                displayError = "Audio format not supported. Please try recording again.";
            }
            else if (errorMessage.Contains("Mindestens ein"))
            {
                displayError = $"Error: {innerMessage ?? errorMessage}";
            }

            return new
            {
                success = false,
                error = displayError,
                errorType = ex.GetType().FullName,
                details = innerMessage ?? errorMessage,
                stackTrace = ex.StackTrace
            };
        }
    }

    private static void EnsureModelLoaded()
    {
        if (processor != null)
        {
            return; // Already loaded
        }

        lock (lockObject)
        {
            if (processor != null)
            {
                return; // Double-check after acquiring lock
            }

            try
            {
                Console.WriteLine("[Whisper.NET] Loading Whisper model...");

                // Get the model file path
                string modelPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Whisper.NET",
                    "ggml-base.bin"
                );

                // Create directory if it doesn't exist
                string modelDir = Path.GetDirectoryName(modelPath);
                if (!Directory.Exists(modelDir))
                {
                    Directory.CreateDirectory(modelDir);
                }

                // Download model if it doesn't exist
                if (!File.Exists(modelPath))
                {
                    Console.WriteLine("[Whisper.NET] Model not found at: " + modelPath);
                    Console.WriteLine("[Whisper.NET] Downloading base model (~75MB)...");
                    Console.WriteLine("[Whisper.NET] This may take a few minutes...");

                    DownloadModelWithRetry(modelPath);

                    if (!File.Exists(modelPath) || new FileInfo(modelPath).Length == 0)
                    {
                        throw new Exception("Model download failed. File is missing or empty.");
                    }

                    Console.WriteLine("[Whisper.NET] Model downloaded successfully!");
                }
                else
                {
                    Console.WriteLine("[Whisper.NET] Model found at: " + modelPath);
                }

                // Load the model
                Console.WriteLine("[Whisper.NET] Creating Whisper factory...");
                whisperFactory = WhisperFactory.FromPath(modelPath);

                Console.WriteLine("[Whisper.NET] Building processor...");
                // Create processor with default settings
                processor = whisperFactory.CreateBuilder()
                    .WithLanguage("auto")
                    .Build();

                Console.WriteLine("[Whisper.NET] Model loaded successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Whisper.NET] Failed to load model: {ex.Message}");
                Console.WriteLine($"[Whisper.NET] Exception type: {ex.GetType().FullName}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"[Whisper.NET] Inner: {ex.InnerException.Message}");
                }
                throw;
            }
        }
    }
}

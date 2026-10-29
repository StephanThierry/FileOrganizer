using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using OpenCvSharp;
using RapidOCRSharpOnnx;
using RapidOCRSharpOnnx.Configurations;
using RapidOCRSharpOnnx.Providers;
using RapidOCRSharpOnnx.Utils;

/// <summary>
/// Cross-platform OCR for video frames using PaddleOCR's PP-OCRv5 models on ONNX Runtime.
/// Replaces Windows.Media.Ocr so the program also runs on Linux and macOS.
///
/// The models (~13 MB) are downloaded once to a "models" folder next to the executable and
/// verified with SHA-256. To run offline, copy the .onnx files into that folder yourself.
/// </summary>
public sealed class VideoTextReader : IDisposable
{
    // Official RapidOCR ONNX exports of PP-OCRv5 (same URLs and checksums as RapidOCR's default_models.yaml).
    private const string ModelBaseUrl = "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.10.0/onnx/PP-OCRv5";

    // Detection: finds the text regions in the frame.
    private static readonly ModelFile DetectorModel = new(
        "ch_PP-OCRv5_det_mobile.onnx", $"{ModelBaseUrl}/det/ch_PP-OCRv5_det_mobile.onnx",
        "4d97c44a20d30a81aad087d6a396b08f786c4635742afc391f6621f5c6ae78ae");

    // Recognition: reads the text in each region. The Latin model covers English plus
    // accented European letters (æ, ø, å, é, ü ...).
    private static readonly ModelFile RecognizerModel = new(
        "latin_PP-OCRv5_rec_mobile.onnx", $"{ModelBaseUrl}/rec/latin_PP-OCRv5_rec_mobile.onnx",
        "b20bd37c168a570f583afbc8cd7925603890efbcdc000a59e22c269d160b5f5a");

    // Lines the recognizer is less sure about than this are discarded (0..1).
    private const float MinLineConfidence = 0.6f;

    private readonly RapidOCRSharp _ocr;

    public VideoTextReader(string modelFolder)
    {
        Directory.CreateDirectory(modelFolder);
        string detPath = EnsureModel(modelFolder, DetectorModel);
        string recPath = EnsureModel(modelFolder, RecognizerModel);

        // No text-line orientation classifier: watermarks and logos in video are upright,
        // so skipping it saves time per frame.
        var config = new OcrConfig(detPath, recPath, LangRec.LATIN, OCRVersion.PPOCRV5);
        config.RecognizerConfig.TextScore = MinLineConfidence;

        // RapidOCRSharpOnnx 1.3.1 uses ProcessorCount / 2 threads, which is 0 on a single-core
        // machine or container. Fail with a clear message instead of an obscure exception.
        if (Environment.ProcessorCount < 2)
            throw new InvalidOperationException(
                "The OCR engine needs at least 2 CPU cores. On a 1-core VM or container, " +
                "start the program with the environment variable DOTNET_PROCESSOR_COUNT=2.");

        _ocr = new RapidOCRSharp(new ExecutionProviderCPU(config));
    }

    /// <summary>Returns the text lines found in a BGR video frame.</summary>
    public List<string> ReadLines(Mat bgrFrame)
    {
        var result = _ocr.RecognizeText(bgrFrame);
        var recognized = result.RecResult?.Data;
        if (recognized == null) return new List<string>();

        return recognized
            .Where(r => r.Score >= MinLineConfidence && !string.IsNullOrWhiteSpace(r.Label))
            .Select(r => r.Label.Trim())
            .ToList();
    }

    public void Dispose() => _ocr.Dispose();

    private static string EnsureModel(string folder, ModelFile model)
    {
        string path = Path.Combine(folder, model.FileName);
        if (File.Exists(path) && Sha256(path) == model.Sha256) return path;

        Console.WriteLine($"[OCR] Downloading model {model.FileName} ...");
        string tempPath = path + ".download";
        using (var http = new HttpClient { Timeout = TimeSpan.FromMinutes(5) })
        {
            // The model host rejects requests without a User-Agent
            http.DefaultRequestHeaders.UserAgent.ParseAdd("FileOrganizer/1.0");
            using var response = http.GetAsync(model.Url, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult();
            response.EnsureSuccessStatusCode();
            using var source = response.Content.ReadAsStream();
            using var target = File.Create(tempPath);
            source.CopyTo(target);
        }

        if (Sha256(tempPath) != model.Sha256)
        {
            File.Delete(tempPath);
            throw new InvalidDataException($"Checksum mismatch for downloaded model {model.FileName}.");
        }

        File.Move(tempPath, path, overwrite: true);
        return path;
    }

    private static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private sealed record ModelFile(string FileName, string Url, string Sha256);
}

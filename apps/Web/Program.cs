using System.Diagnostics;
using System.Net;
using System.Security.Cryptography;
using FileOrganizer;
using FileOrganizer.Web.Components;
using Microsoft.AspNetCore.DataProtection;

// Usage: FileOrganizer.Web [--data <folder>] [--urls http://0.0.0.0:5078] [--token <secret>] [--no-browser]
//   --data   folder with settings.json                         (or FILEORG_DATA; default: the current folder if it holds
//                                                                settings.json, else the user's application data folder)
//   --urls   where to listen                                    (or FILEORG_URLS; default http://127.0.0.1:5078, this computer only)
//   --token  needed to open the page when listening on the network (or FILEORG_TOKEN; "off" turns it off).
//            Without one, a token is created and kept in <data>/.token, so it stays the same after restarts.

string? Arg(string name) => DataLocation.Arg(args, name);
string? Env(string name) => string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(name)) ? null : Environment.GetEnvironmentVariable(name);

string dataFolder = DataLocation.ResolveDataFolder(args, null) ?? DataLocation.DefaultDataFolder();
string urls = Arg("--urls") ?? Env("FILEORG_URLS") ?? "http://127.0.0.1:5078";
bool openBrowser = !args.Contains("--no-browser");

// The content root is the app folder (static files)
var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args = args.Where(a => a != "--no-browser").ToArray(),
    ContentRootPath = AppContext.BaseDirectory
});
builder.WebHost.UseUrls(urls);
// Serve the framework files straight from the build output too (dotnet run), not only from a publish
builder.WebHost.UseStaticWebAssets();

var controller = AppController.Open(dataFolder);
builder.Services.AddSingleton(controller);
// Keep the keys that protect forms in the data folder, so a restart doesn't break open pages
builder.Services.AddDataProtection()
    .PersistKeysToFileSystem(new DirectoryInfo(Path.Combine(controller.DataFolder, ".keys")))
    .SetApplicationName("FileOrganizer");
builder.Services.AddRazorComponents().AddInteractiveServerComponents();

// Listening beyond this computer: the page can move files, so require a token
bool localOnly = urls.Split(';').All(u => u.Contains("127.0.0.1") || u.Contains("localhost") || u.Contains("[::1]"));
string? token = Arg("--token") ?? Env("FILEORG_TOKEN") ?? (localOnly ? null : StoredToken(controller.DataFolder));
if (string.Equals(token, "off", StringComparison.OrdinalIgnoreCase)) token = null;

// A token that survives restarts: created once, kept in the data folder
static string StoredToken(string folder)
{
    string file = Path.Combine(folder, ".token");
    if (File.Exists(file)) { var t = File.ReadAllText(file).Trim(); if (t.Length >= 8) return t; }
    string token = Convert.ToHexString(RandomNumberGenerator.GetBytes(12)).ToLowerInvariant();
    File.WriteAllText(file, token);
    return token;
}

var app = builder.Build();
if (!app.Environment.IsDevelopment()) app.UseExceptionHandler("/Error", createScopeForErrors: true);

if (token != null)
{
    const string cookie = "fo_token";
    app.Use(async (ctx, next) =>
    {
        if (ctx.Request.Query["token"] == token)
        {
            ctx.Response.Cookies.Append(cookie, token, new CookieOptions { HttpOnly = true, SameSite = SameSiteMode.Strict });
            ctx.Response.Redirect(ctx.Request.Path);
            return;
        }
        if (ctx.Request.Cookies[cookie] != token)
        {
            ctx.Response.StatusCode = 401;
            ctx.Response.ContentType = "text/plain";
            await ctx.Response.WriteAsync("Open the address with ?token=... that the program printed when it started.");
            return;
        }
        await next();
    });
}

app.UseAntiforgery();
app.MapStaticAssets();

// A picture of a video of the current result: /thumb/12 (12 = its place in the result), or /thumb/12?part=0..2 for one of the three
// spread over the video. Made on the first request.
app.MapGet("/thumb/{n:int}", (int n, int? part) =>
{
    var d = controller.Result?.Decisions.ElementAtOrDefault(n);
    var path = d == null ? null : part is { } p ? controller.ThumbOf(d, p) : controller.ThumbOf(d);
    return path != null && File.Exists(path) ? Results.File(path, "image/png", enableRangeProcessing: false) : Results.NotFound();
});

// A picture made during a scan, by its file name: /thumbfile/<hash>.png
var thumbFile = new System.Text.RegularExpressions.Regex(@"^[0-9a-f]{40}\.png$");
app.MapGet("/thumbfile/{name}", (string name) =>
{
    if (!thumbFile.IsMatch(name)) return Results.NotFound();
    var path = Path.Combine(controller.ThumbFolder, name);
    return File.Exists(path) ? Results.File(path, "image/png", enableRangeProcessing: false) : Results.NotFound();
});

// A file of the current result, for the player: /media/12. Only files of the result are served.
var mediaTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
{
    [".mp4"] = "video/mp4", [".m4v"] = "video/mp4", [".mov"] = "video/quicktime", [".webm"] = "video/webm",
    [".mkv"] = "video/x-matroska", [".ogv"] = "video/ogg", [".avi"] = "video/x-msvideo", [".wmv"] = "video/x-ms-wmv",
    [".flv"] = "video/x-flv", [".mpg"] = "video/mpeg", [".mpeg"] = "video/mpeg",
};
app.MapGet("/media/{n:int}", (int n) =>
{
    var d = controller.Result?.Decisions.ElementAtOrDefault(n);
    var path = d == null ? null : controller.PathOf(d);
    if (path == null || !File.Exists(path)) return Results.NotFound();
    var type = mediaTypes.TryGetValue(Path.GetExtension(path), out var t) ? t : "application/octet-stream";
    return Results.File(path, type, enableRangeProcessing: true);
});

app.MapRazorComponents<App>().AddInteractiveServerRenderMode();

string address = urls.Split(';')[0].Replace("0.0.0.0", Dns.GetHostName()).Replace("[::]", Dns.GetHostName());
string openUrl = token == null ? address : $"{address}/?token={token}";
app.Lifetime.ApplicationStopping.Register(() => { controller.StopScan(); controller.Dispose(); });
app.Lifetime.ApplicationStarted.Register(() =>
{
    Console.WriteLine();
    Console.WriteLine($"File Organizer is running. Data folder: {controller.DataFolder}");
    Console.WriteLine($"Open: {openUrl}");
    Console.WriteLine("Press Ctrl+C to stop.");
    if (openBrowser && localOnly)
    {
        try { Process.Start(new ProcessStartInfo(openUrl) { UseShellExecute = true }); }
        catch { /* no browser on this machine (server) - the address is printed above */ }
    }
});

app.Run();

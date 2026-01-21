using System.Diagnostics;
using System.Globalization;
using System.Net.Sockets;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.WindowsServices;
using System.Drawing.Printing;

namespace BlPrintService;

public static class Program
{
    // ===================== CONFIG (STATIC PATHS) =====================
    private const string TEMPLATE_PATH = @"C:\Users\YasserBOUAABANE\Downloads\Captures Slv\Mapping delivery note -Alexsys.xlsx";
    private const string OUT_DIR = @"C:\out";
    private const string SOFFICE_PATH = @"C:\Program Files\LibreOffice\program\soffice.exe";

    // Ghostscript console exe (only needed for mode=pcl)
    private const string GHOSTSCRIPT_EXE = @"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe";

    // Sumatra (used for mode=driver)
    private const string SUMATRA_EXE = @"C:\Users\YasserBOUAABANE\AppData\Local\SumatraPDF\SumatraPDF.exe";

    // Raw socket printing (mode=pcl)
    private const string PRINTER_IP = "10.8.197.25";
    private const int PRINTER_PORT = 9100;

    // Konica driver printer name (mode=driver)
    private const string PRINTER_NAME_CONTAINS = "KONICA MINOLTA bizhub 5000i";

    public static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);

        // Windows Service support (NuGet: Microsoft.Extensions.Hosting.WindowsServices)
        builder.Host.UseWindowsService(options => options.ServiceName = "BL Print Service");

        builder.Services.ConfigureHttpJsonOptions(o =>
        {
            o.SerializerOptions.PropertyNameCaseInsensitive = true;
        });

        var app = builder.Build();

        // POST /bl/print?mode=driver|pcl&page=2
        app.MapPost("/bl/print", async (
            [FromBody] BlJson payload,
            [FromQuery] string? mode,
            [FromQuery] int? page,
            CancellationToken ct) =>
        {
            mode ??= "driver";
            int targetPage = page ?? 2;

            Directory.CreateDirectory(OUT_DIR);

            // Validate required files
            if (!File.Exists(TEMPLATE_PATH)) return Results.Problem($"Template not found: {TEMPLATE_PATH}");
            if (!File.Exists(SOFFICE_PATH)) return Results.Problem($"LibreOffice not found: {SOFFICE_PATH}");

            if (mode.Equals("driver", StringComparison.OrdinalIgnoreCase))
            {
                if (!File.Exists(SUMATRA_EXE)) return Results.Problem($"SumatraPDF not found: {SUMATRA_EXE}");
            }
            else if (mode.Equals("pcl", StringComparison.OrdinalIgnoreCase))
            {
                if (!File.Exists(GHOSTSCRIPT_EXE)) return Results.Problem($"Ghostscript not found: {GHOSTSCRIPT_EXE}");
            }
            else
            {
                return Results.Problem("Invalid mode. Use mode=driver or mode=pcl");
            }

            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            var xlsxPath = Path.Combine(OUT_DIR, $"BL_{stamp}.xlsx");

            // 1) Fill Excel
            ExcelBlGenerator.FillExcel(TEMPLATE_PATH, xlsxPath, payload);

            // 2) XLSX -> PDF (LibreOffice headless)
            var pdfPath = await PdfExporter.ExportPdfWithLibreOfficeAsync(SOFFICE_PATH, xlsxPath, OUT_DIR, ct);

            // 3) Print
            if (mode.Equals("driver", StringComparison.OrdinalIgnoreCase))
            {
                var printerName = FindPrinterNameOrThrow(PRINTER_NAME_CONTAINS);

                // Sumatra printing options:
                // -print-to "Printer"
                // -print-settings "2,fit,paper=A5"
                // -silent -exit-when-done "file.pdf"
                // (Supported by Sumatra command line docs) :contentReference[oaicite:3]{index=3}
                await SumatraPrinter.PrintPdfPageA5Async(
                    sumatraExe: SUMATRA_EXE,
                    printerName: printerName,
                    pdfPath: pdfPath,
                    page: targetPage,
                    ct: ct
                );

                return Results.Ok(new
                {
                    mode = "driver",
                    page = targetPage,
                    printerName,
                    xlsxPath,
                    pdfPath
                });
            }
            else
            {
                // mode=pcl: PDF -> PCL6 (PCL-XL) via Ghostscript pxlmono :contentReference[oaicite:4]{index=4}
                var pclPath = Path.Combine(OUT_DIR, $"BL_{stamp}_p{targetPage}_A5.pcl");

                await GhostscriptPcl.ConvertPdfToPclXlAsync(
                    gsExe: GHOSTSCRIPT_EXE,
                    pdfPath: pdfPath,
                    pclOutPath: pclPath,
                    firstPage: targetPage,
                    lastPage: targetPage,
                    ct: ct
                );

                var pclSize = new FileInfo(pclPath).Length;
                if (pclSize < 500) return Results.Problem($"PCL too small ({pclSize} bytes). PDF page {targetPage} may not exist.");

                await Raw9100Printer.SendRawFileAsync(PRINTER_IP, PRINTER_PORT, pclPath, ct);

                return Results.Ok(new
                {
                    mode = "pcl",
                    page = targetPage,
                    printedTo = $"{PRINTER_IP}:{PRINTER_PORT}",
                    xlsxPath,
                    pdfPath,
                    pclPath,
                    pclBytes = pclSize
                });
            }
        });

        app.Run();
    }

    private static string FindPrinterNameOrThrow(string contains)
    {
        foreach (string p in PrinterSettings.InstalledPrinters)
        {
            if (p.Contains(contains, StringComparison.OrdinalIgnoreCase))
                return p;
        }

        throw new InvalidOperationException(
            $"Printer not found (contains='{contains}'). Check PrinterSettings.InstalledPrinters."
        );
    }
}

// ====================== MODELS ======================
public sealed class BlJson
{
    public string? Site { get; set; }
    public string? BonDeLivraison { get; set; }
    public ClientJson? Client { get; set; }
    public TransportJson? Transport { get; set; }
    public PesageJson? Pesage { get; set; }
    public List<ProductJson>? Produits { get; set; }
}
public sealed class ClientJson
{
    public string? CodeSap { get; set; }
    public string? Name { get; set; }
    public string? Chantier { get; set; }
    public string? BonDeCommande { get; set; }
}
public sealed class TransportJson
{
    public string? Transporteur { get; set; }
    public string? Matricule { get; set; }
    public string? Chauffeur { get; set; }
    public List<string>? Scelles { get; set; }
}
public sealed class PesageJson
{
    public int? PoidsVide { get; set; }
    public int? PoidsBrut { get; set; }
    public DateTime? PabEntryAt { get; set; }
    public DateTime? PabExitAt { get; set; }
}
public sealed class ProductJson
{
    public string? Code { get; set; }
    public string? Libelle { get; set; }
    public decimal? Quantite { get; set; }
    public int? Sacs { get; set; }
}

// ====================== EXCEL FILL ======================
public static class ExcelBlGenerator
{
    public static void FillExcel(string templatePath, string outputPath, BlJson p)
    {
        if (File.Exists(outputPath)) File.Delete(outputPath);

        using var wb = new XLWorkbook(templatePath);
        var ws = wb.Worksheet("BL");

        if (ws.Protection.IsProtected) ws.Unprotect();

        var entry = p.Pesage?.PabEntryAt;
        var exit = p.Pesage?.PabExitAt;
        DateTime dateP = entry ?? DateTime.Now;
        DateTime dateL = exit ?? dateP;

        // Header
        SetMerged(ws, "E1:E1", $"Date:{dateP:dd.MM.yyyy}");
        SetMerged(ws, "E3:F3", $"Site : {(string.IsNullOrWhiteSpace(p.Site) ? "Asment de Témara" : p.Site)}");
        SetMerged(ws, "G4:H4", $"BON DE LIVRAISON {p.BonDeLivraison ?? ""}");

        // Client
        SetMerged(ws, "F9:G9", p.Client?.CodeSap ?? "");
        SetMerged(ws, "F10:G10", p.Client?.Name ?? "");
        SetMerged(ws, "F11:G11", p.Client?.BonDeCommande ?? "");
        SetMerged(ws, "E12:G13", p.Client?.Chantier ?? "");

        // Transport
        SetMerged(ws, "I9:J9", p.Transport?.Transporteur ?? "");
        SetMerged(ws, "I10:J10", p.Transport?.Matricule ?? "");
        SetMerged(ws, "I11:J11", p.Transport?.Chauffeur ?? "");

        var seals = p.Transport?.Scelles ?? new List<string>();
        var seal1 = seals.Count > 0 ? seals[0] : "";
        var seal2 = seals.Count > 1 ? string.Join(",", seals.Skip(1)) : "";
        SetMerged(ws, "I12:J12", seal1);
        SetMerged(ws, "I13:J13", seal2);

        // Pesage
        SetMerged(ws, "F21:G21", p.Pesage?.PoidsVide?.ToString() ?? "");
        SetMerged(ws, "F22:G22", p.Pesage?.PoidsBrut?.ToString() ?? "");
        SetMerged(ws, "I21:J21", entry.HasValue ? entry.Value.ToString("ddMMyyyy HHmmss", CultureInfo.InvariantCulture) : "");
        SetMerged(ws, "I22:J22", exit.HasValue ? exit.Value.ToString("ddMMyyyy HHmmss", CultureInfo.InvariantCulture) : "");
        SetMerged(ws, "F23:G23", dateP.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture));
        SetMerged(ws, "I23:J23", dateL.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture));

        // Products (max 2)
        var prods = p.Produits ?? new List<ProductJson>();

        // Product 1
        if (prods.Count >= 1)
        {
            var pr1 = prods[0];
            SetMerged(ws, "E16:E17", pr1.Code ?? "");
            SetMerged(ws, "F16:H17", pr1.Libelle ?? "");
            SetMerged(ws, "I16:I17", FormatQty(pr1.Quantite));
            SetMerged(ws, "J16:J17", pr1.Sacs?.ToString() ?? "");
        }
        else
        {
            ClearMerged(ws, "E16:E17");
            ClearMerged(ws, "F16:H17");
            ClearMerged(ws, "I16:I17");
            ClearMerged(ws, "J16:J17");
        }

        // Product 2
        if (prods.Count >= 2)
        {
            var pr2 = prods[1];
            SetMerged(ws, "E18:E18", pr2.Code ?? "");
            SetMerged(ws, "F18:H18", pr2.Libelle ?? "");
            SetMerged(ws, "I18:I18", FormatQty(pr2.Quantite));
            SetMerged(ws, "J18:J18", pr2.Sacs?.ToString() ?? "");
        }
        else
        {
            ClearMerged(ws, "E18:E18");
            ClearMerged(ws, "F18:H18");
            ClearMerged(ws, "I18:I18");
            ClearMerged(ws, "J18:J18");
        }

        wb.SaveAs(outputPath);
    }

    private static void SetMerged(IXLWorksheet ws, string rangeAddress, string value)
    {
        var rng = ws.Range(rangeAddress);
        rng.Clear(XLClearOptions.Contents);

        // Write into the top-left of the merged range
        var cell = rng.FirstCell();
        var target = cell.IsMerged() ? cell.MergedRange().FirstCell() : cell;

        target.FormulaA1 = "";
        target.SetValue(value);
    }

    private static void ClearMerged(IXLWorksheet ws, string rangeAddress)
        => ws.Range(rangeAddress).Clear(XLClearOptions.Contents);

    private static string FormatQty(decimal? q)
    {
        if (q is null) return "";
        return (decimal.Truncate(q.Value) == q.Value)
            ? ((int)q.Value).ToString(CultureInfo.InvariantCulture)
            : q.Value.ToString(CultureInfo.InvariantCulture);
    }
}

// ====================== XLSX -> PDF (LibreOffice) ======================
// LibreOffice supports: --headless --convert-to pdf --outdir ... :contentReference[oaicite:5]{index=5}
public static class PdfExporter
{
    public static async Task<string> ExportPdfWithLibreOfficeAsync(
        string sofficePath,
        string xlsxPath,
        string outDir,
        CancellationToken ct)
    {
        Directory.CreateDirectory(outDir);

        var psi = new ProcessStartInfo
        {
            FileName = sofficePath,
            Arguments = $"--headless --nologo --nofirststartwizard --norestore --convert-to pdf --outdir \"{outDir}\" \"{xlsxPath}\"",
            UseShellExecute = false,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            CreateNoWindow = true
        };

        using var p = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start LibreOffice.");
        var stderr = await p.StandardError.ReadToEndAsync();
        await p.WaitForExitAsync(ct);

        if (p.ExitCode != 0) throw new Exception($"LibreOffice export failed: {stderr}");

        var pdfPath = Path.Combine(outDir, Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf");
        if (!File.Exists(pdfPath)) throw new FileNotFoundException("PDF not generated", pdfPath);

        await WaitForFileReadyAsync(pdfPath, timeoutMs: 8000, ct);
        return pdfPath;
    }

    private static async Task WaitForFileReadyAsync(string path, int timeoutMs, CancellationToken ct)
    {
        var sw = Stopwatch.StartNew();
        while (true)
        {
            ct.ThrowIfCancellationRequested();

            try
            {
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
                return;
            }
            catch (IOException)
            {
                if (sw.ElapsedMilliseconds > timeoutMs) throw;
                await Task.Delay(200, ct);
            }
        }
    }
}

// ====================== PDF -> PCL6 (Ghostscript) ======================
// pxlmono outputs HP PCL-XL (PCL6). :contentReference[oaicite:6]{index=6}
public static class GhostscriptPcl
{
    public static async Task ConvertPdfToPclXlAsync(
        string gsExe,
        string pdfPath,
        string pclOutPath,
        int firstPage,
        int lastPage,
        CancellationToken ct)
    {
        if (File.Exists(pclOutPath)) File.Delete(pclOutPath);

        var args =
            $"-dSAFER -dBATCH -dNOPAUSE " +
            $"-dFirstPage={firstPage} -dLastPage={lastPage} " +
            $"-sDEVICE=pxlmono -sPAPERSIZE=a5 -dFIXEDMEDIA " +
            $"-sOutputFile=\"{pclOutPath}\" \"{pdfPath}\"";

        var psi = new ProcessStartInfo
        {
            FileName = gsExe,
            Arguments = args,
            UseShellExecute = false,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            CreateNoWindow = true
        };

        using var p = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start Ghostscript.");
        var stdout = await p.StandardOutput.ReadToEndAsync();
        var stderr = await p.StandardError.ReadToEndAsync();
        await p.WaitForExitAsync(ct);

        if (p.ExitCode != 0)
            throw new Exception($"Ghostscript PDF->PCL-XL failed (code {p.ExitCode}).\nSTDOUT:\n{stdout}\nSTDERR:\n{stderr}");

        if (!File.Exists(pclOutPath))
            throw new FileNotFoundException("PCL not generated", pclOutPath);
    }
}

// ====================== PRINTING (Sumatra via Driver) ======================
// Sumatra supports -print-to and -print-settings with page ranges + paper=A5 + fit. :contentReference[oaicite:7]{index=7}
public static class SumatraPrinter
{
    public static async Task PrintPdfPageA5Async(
        string sumatraExe,
        string printerName,
        string pdfPath,
        int page,
        CancellationToken ct)
    {
        var psi = new ProcessStartInfo
        {
            FileName = sumatraExe,
            Arguments =
                $"-print-to \"{printerName}\" " +
                $"-print-settings \"{page},fit,paper=A5\" " +
                "-silent -exit-when-done " +
                $"\"{pdfPath}\"",
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var p = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start SumatraPDF.");
        await p.WaitForExitAsync(ct);

        if (p.ExitCode != 0)
            throw new Exception($"SumatraPDF printing failed. ExitCode={p.ExitCode}");
    }
}

// ====================== RAW 9100 SEND (PCL) ======================
public static class Raw9100Printer
{
    public static async Task SendRawFileAsync(string ip, int port, string path, CancellationToken ct)
    {
        using var tcp = new TcpClient();
        tcp.NoDelay = true;
        await tcp.ConnectAsync(ip, port, ct);

        await using var net = tcp.GetStream();
        await using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

        await fs.CopyToAsync(net, ct);

        // Form feed is often used as end-of-job marker for PCL streams
        await net.WriteAsync(new byte[] { 0x0C }, ct);
        await net.FlushAsync(ct);
    }
}

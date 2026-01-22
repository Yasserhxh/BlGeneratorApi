using System.Diagnostics;
using System.Globalization;
using System.Net.Sockets;
using System.Drawing.Printing;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.WindowsServices;

namespace BlPrintService;

public static class Program
{
    // ===================== CONFIG =====================
    private const string TEMPLATE_PATH = @"C:\Template.xlsx";
    private const string OUT_DIR = @"C:\out";
    private const string SOFFICE_PATH = @"C:\Program Files\LibreOffice\program\soffice.exe";
    private const string GHOSTSCRIPT_EXE = @"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe";
    private const string SUMATRA_EXE = @"C:\Users\cachapuz\AppData\Local\SumatraPDF\SumatraPDF.exe";

    private const string PRINTER_IP = "10.8.197.25";
    private const int PRINTER_PORT = 9100;
    private const string PRINTER_NAME_CONTAINS = "KONICA MINOLTA bizhub 5000i";

    private static readonly string[] COPY_LABELS =
    {
        "Client",
        "Pont Bascule",
        "Expedition",
        "Commercial",
        "Souche"
    };

    public static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);
        builder.Host.UseWindowsService(o => o.ServiceName = "BL Print Service");
        var app = builder.Build();

        app.MapPost("/bl/print", async (
            [FromBody] BlJson payload,
            [FromQuery] string? mode,
            [FromQuery] int? page,
            CancellationToken ct) =>
        {
            mode ??= "driver";
            int targetPage = page ?? 2;

            Directory.CreateDirectory(OUT_DIR);

            string? printerName = null;
            if (mode == "driver")
                printerName = FindPrinterNameOrThrow(PRINTER_NAME_CONTAINS);

            foreach (var label in COPY_LABELS)
            {
                var stamp = $"{DateTime.Now:yyyyMMdd_HHmmss}_{label.Replace(" ", "_")}";
                var xlsxPath = Path.Combine(OUT_DIR, $"BL_{stamp}.xlsx");

                // 1) Fill Excel (E28 changes here)
                ExcelBlGenerator.FillExcel(
                    TEMPLATE_PATH,
                    xlsxPath,
                    payload,
                    label
                );

                // 2) Excel -> PDF
                var pdfPath = await PdfExporter.ExportPdfWithLibreOfficeAsync(
                    SOFFICE_PATH,
                    xlsxPath,
                    OUT_DIR,
                    ct
                );

                // 3) Print
                if (mode == "driver")
                {
                    await SumatraPrinter.PrintPdfPageA5Async(
                        SUMATRA_EXE,
                        printerName!,
                        pdfPath,
                        targetPage,
                        ct
                    );
                }
                else
                {
                    var pclPath = Path.Combine(
                        OUT_DIR,
                        $"BL_{stamp}_p{targetPage}.pcl"
                    );

                    await GhostscriptPcl.ConvertPdfToPclXlAsync(
                        GHOSTSCRIPT_EXE,
                        pdfPath,
                        pclPath,
                        targetPage,
                        targetPage,
                        ct
                    );

                    await Raw9100Printer.SendRawFileAsync(
                        PRINTER_IP,
                        PRINTER_PORT,
                        pclPath,
                        ct
                    );
                }
            }

            return Results.Ok(new
            {
                copies = COPY_LABELS.Length,
                labels = COPY_LABELS,
                mode,
                page = targetPage
            });
        });

        app.Run();
    }

    private static string FindPrinterNameOrThrow(string contains)
    {
        foreach (string p in PrinterSettings.InstalledPrinters)
            if (p.Contains(contains, StringComparison.OrdinalIgnoreCase))
                return p;

        throw new InvalidOperationException("Printer not found");
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

// ====================== EXCEL ======================
public static class ExcelBlGenerator
{
    public static void FillExcel(
        string templatePath,
        string outputPath,
        BlJson p,
        string copyLabel)
    {
        if (File.Exists(outputPath)) File.Delete(outputPath);

        using var wb = new XLWorkbook(templatePath);
        var ws = wb.Worksheet("BL");

        if (ws.Protection.IsProtected)
            ws.Unprotect();

        // ===== COPY LABEL =====
        ws.Range("E28:E28").Clear(XLClearOptions.Contents);
        ws.Cell("E28").SetValue(copyLabel);

        // ===== HEADER =====
        ws.Cell("G4").SetValue($"BON DE LIVRAISON {p.BonDeLivraison ?? ""}");
        ws.Cell("F9").SetValue(p.Client?.CodeSap ?? "");
        ws.Cell("F10").SetValue(p.Client?.Name ?? "");
        ws.Cell("F11").SetValue(p.Client?.BonDeCommande ?? "");
        ws.Cell("E12").SetValue(p.Client?.Chantier ?? "");

        ws.Cell("I9").SetValue(p.Transport?.Transporteur ?? "");
        ws.Cell("I10").SetValue(p.Transport?.Matricule ?? "");
        ws.Cell("I11").SetValue(p.Transport?.Chauffeur ?? "");

        wb.SaveAs(outputPath);
    }
}

// ====================== PDF EXPORT ======================
public static class PdfExporter
{
    public static async Task<string> ExportPdfWithLibreOfficeAsync(
        string sofficePath,
        string xlsxPath,
        string outDir,
        CancellationToken ct)
    {
        var psi = new ProcessStartInfo
        {
            FileName = sofficePath,
            Arguments = $"--headless --convert-to pdf --outdir \"{outDir}\" \"{xlsxPath}\"",
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var p = Process.Start(psi)!;
        await p.WaitForExitAsync(ct);

        var pdfPath = Path.Combine(
            outDir,
            Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf"
        );

        return pdfPath;
    }
}

// ====================== GHOSTSCRIPT ======================
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
        var args =
            $"-dBATCH -dNOPAUSE -sDEVICE=pxlmono -sPAPERSIZE=a5 " +
            $"-dFirstPage={firstPage} -dLastPage={lastPage} " +
            $"-sOutputFile=\"{pclOutPath}\" \"{pdfPath}\"";

        var psi = new ProcessStartInfo
        {
            FileName = gsExe,
            Arguments = args,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var p = Process.Start(psi)!;
        await p.WaitForExitAsync(ct);
    }
}

// ====================== SUMATRA ======================
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

        using var p = Process.Start(psi)!;
        await p.WaitForExitAsync(ct);
    }
}

// ====================== RAW 9100 ======================
public static class Raw9100Printer
{
    public static async Task SendRawFileAsync(
        string ip,
        int port,
        string path,
        CancellationToken ct)
    {
        using var tcp = new TcpClient();
        await tcp.ConnectAsync(ip, port, ct);

        await using var net = tcp.GetStream();
        await using var fs = File.OpenRead(path);

        await fs.CopyToAsync(net, ct);
        await net.WriteAsync(new byte[] { 0x0C }, ct); // Form Feed
        await net.FlushAsync(ct);
    }
}

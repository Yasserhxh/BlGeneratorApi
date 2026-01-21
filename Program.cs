using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace BlGeneratorApi;

public sealed class Program
{
    public static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);

        builder.Services.ConfigureHttpJsonOptions(o =>
        {
            o.SerializerOptions.PropertyNameCaseInsensitive = true;
        });

        var app = builder.Build();

        app.MapPost("/bl/generate", async ([FromBody] BlJson payload) =>
        {
            // ===== STATIC PATHS (change these) =====
            const string TEMPLATE_PATH = @"C:\Users\YasserBOUAABANE\Downloads\Captures Slv\Mapping delivery note -Alexsys.xlsx";
            const string OUT_DIR = @"C:\out";
            const string SOFFICE_PATH = @"C:\Program Files\LibreOffice\program\soffice.exe";

            if (!File.Exists(TEMPLATE_PATH))
                return Results.Problem($"Template not found: {TEMPLATE_PATH}");

            /*if (!File.Exists(SOFFICE_PATH))
                return Results.Problem($"LibreOffice not found: {SOFFICE_PATH}");*/

            Directory.CreateDirectory(OUT_DIR);

            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            var xlsxPath = Path.Combine(OUT_DIR, $"BL_{stamp}.xlsx");

            ExcelBlGenerator.FillExcel(TEMPLATE_PATH, xlsxPath, payload);
            //var pdfPath = await PdfExporter.ExportPdfWithLibreOfficeAsync(SOFFICE_PATH, xlsxPath, OUT_DIR);

            return Results.Ok(new { xlsxPath/*, pdfPath*/ });
        })
        .WithName("GenerateBL")
        .Accepts<BlJson>("application/json")
        .Produces(200)
        .Produces(500);

        app.Run();
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


// ====================== EXCEL GENERATOR ======================
public static class ExcelBlGenerator
{
    public static void FillExcel(string templatePath, string outputPath, BlJson p)
    {
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        using var wb = new XLWorkbook(templatePath);

        var ws = wb.Worksheet("BL");

        // If the sheet is protected, nothing will change
        if (ws.Protection.IsProtected)
            ws.Unprotect();

        var entry = p.Pesage?.PabEntryAt;
        var exit = p.Pesage?.PabExitAt;

        DateTime dateP = entry ?? DateTime.Now;
        DateTime dateL = exit ?? dateP;

        // Header
        SetMerged(ws, "E1:E1", $"Date:{DateDot(dateP)}");
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

        // Scellés (2 lines)
        var seals = p.Transport?.Scelles ?? new List<string>();
        string sealLine1 = seals.Count > 0 ? seals[0] : "";
        string sealLine2 = seals.Count > 1 ? seals[1] : "";
        if (seals.Count > 2) sealLine2 = string.Join(",", seals.Skip(1));

        SetMerged(ws, "I12:J12", sealLine1);
        SetMerged(ws, "I13:J13", sealLine2);

        // Pesage
        SetMerged(ws, "F21:G21", p.Pesage?.PoidsVide ?? 0);
        SetMerged(ws, "F22:G22", p.Pesage?.PoidsBrut ?? 0);

        SetMerged(ws, "I21:J21", entry.HasValue ? DateTimeCompact(entry.Value) : "");
        SetMerged(ws, "I22:J22", exit.HasValue ? DateTimeCompact(exit.Value) : "");

        SetMerged(ws, "F23:G23", DateDot(dateP));
        SetMerged(ws, "I23:J23", DateDot(dateL));

        // Products (max 2)
        var produits = p.Produits ?? new List<ProductJson>();

        if (produits.Count >= 1)
        {
            var pr1 = produits[0];
            SetMerged(ws, "E16:E17", pr1.Code ?? "");
            SetMerged(ws, "F16:H17", pr1.Libelle ?? "");
            SetMerged(ws, "I16:I17", FormatQty(pr1.Quantite));
            SetMerged(ws, "J16:J17", pr1.Sacs ?? 0);
        }
        else
        {
            ClearRange(ws, "E16:E17"); ClearRange(ws, "F16:H17"); ClearRange(ws, "I16:I17"); ClearRange(ws, "J16:J17");
        }

        if (produits.Count >= 2)
        {
            var pr2 = produits[1];
            SetMerged(ws, "E18:E18", pr2.Code ?? "");
            SetMerged(ws, "F18:H18", pr2.Libelle ?? "");
            SetMerged(ws, "I18:I18", FormatQty(pr2.Quantite));
            SetMerged(ws, "J18:J18", pr2.Sacs ?? 0);
        }
        else
        {
            ClearRange(ws, "E18:E18"); ClearRange(ws, "F18:H18"); ClearRange(ws, "I18:I18"); ClearRange(ws, "J18:J18");
        }

        wb.SaveAs(outputPath);
    }

    private static void ClearRange(IXLWorksheet ws, string rangeAddress)
    {
        var rng = ws.Range(rangeAddress);
        // Clear displayed content in the target merged cell (or normal range)
        var first = rng.FirstCell();
        var target = first.IsMerged() ? first.MergedRange().FirstCell() : first;
        target.Clear(XLClearOptions.Contents);
        target.FormulaA1 = string.Empty;
    }

    private static void SetMerged(IXLWorksheet ws, string rangeAddress, object? value)
    {
        var rng = ws.Range(rangeAddress);
        var first = rng.FirstCell();
        var target = first.IsMerged() ? first.MergedRange().FirstCell() : first;

        // clear contents only (keep styles), and remove formula if exists
        target.Clear(XLClearOptions.Contents);
        target.FormulaA1 = string.Empty;

        if (value is null) return;

        switch (value)
        {
            case string s: target.SetValue(s); break;
            case int i: target.SetValue(i); break;
            case long l: target.SetValue(l); break;
            case double d: target.SetValue(d); break;
            case float f: target.SetValue((double)f); break;
            case decimal m: target.SetValue((double)m); break;
            case DateTime dt: target.SetValue(dt); break;
            case bool b: target.SetValue(b); break;
            default: target.SetValue(value.ToString() ?? ""); break;
        }
    }

    private static string DateDot(DateTime dt) => dt.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
    private static string DateTimeCompact(DateTime dt) => dt.ToString("ddMMyyyy HHmmss", CultureInfo.InvariantCulture);

    private static string FormatQty(decimal? q)
    {
        if (q == null) return "";
        if (decimal.Truncate(q.Value) == q.Value) return ((int)q.Value).ToString(CultureInfo.InvariantCulture);
        return q.Value.ToString(CultureInfo.InvariantCulture);
    }
}


// ====================== PDF EXPORT ======================
public static class PdfExporter
{
    public static async Task<string> ExportPdfWithLibreOfficeAsync(string sofficePath, string xlsxPath, string outDir)
    {
        Directory.CreateDirectory(outDir);

        var psi = new ProcessStartInfo
        {
            FileName = sofficePath,
            Arguments =
                $"--headless --nologo --nofirststartwizard --norestore " +
                $"--convert-to pdf --outdir \"{outDir}\" \"{xlsxPath}\"",
            UseShellExecute = false,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            CreateNoWindow = true
        };

        using var p = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start LibreOffice");
        var stdout = await p.StandardOutput.ReadToEndAsync();
        var stderr = await p.StandardError.ReadToEndAsync();
        await p.WaitForExitAsync();

        if (p.ExitCode != 0)
            throw new Exception($"LibreOffice export failed (code {p.ExitCode}).\nSTDOUT:\n{stdout}\nSTDERR:\n{stderr}");

        var pdfPath = Path.Combine(outDir, Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf");
        if (!File.Exists(pdfPath)) throw new FileNotFoundException("PDF not generated", pdfPath);

        return pdfPath;
    }
}

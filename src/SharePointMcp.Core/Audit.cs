using System.Text.Json;
using System.Text.Json.Serialization;

namespace SharePointMcp.Core;

public sealed record AuditRecord(
    DateTimeOffset Timestamp,
    string Operation,
    string CanonicalSiteIdentifier,
    string AffectedObject,
    string Outcome,
    string? CorrelationId,
    bool DryRun = false);

public interface IAuditLog
{
    Task WriteAsync(AuditRecord record, CancellationToken cancellationToken = default);
}

public sealed class JsonLinesAuditLog(string directory) : IAuditLog
{
    private readonly SemaphoreSlim _gate = new(1, 1);
    private static readonly JsonSerializerOptions CompactJson = new(JsonDefaults.Options) { WriteIndented = false };

    public async Task WriteAsync(AuditRecord record, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, $"mutations-{record.Timestamp:yyyy-MM}.jsonl");
        var safe = record with
        {
            Operation = Redact(record.Operation),
            AffectedObject = Redact(record.AffectedObject),
            Outcome = Redact(record.Outcome),
            CorrelationId = record.CorrelationId is null ? null : Redact(record.CorrelationId)
        };
        var line = JsonSerializer.Serialize(safe, CompactJson) + Environment.NewLine;
        await _gate.WaitAsync(cancellationToken);
        try
        {
            await File.AppendAllTextAsync(path, line, cancellationToken);
        }
        finally { _gate.Release(); }
    }

    private static string Redact(string value)
    {
        if (new[] { "Bearer ", "refresh_token", "access_token", "client_secret", "password", "cookie" }
            .Any(marker => value.Contains(marker, StringComparison.OrdinalIgnoreCase)))
            return "[REDACTED]";
        return value.Length <= 2048 ? value : value[..2048];
    }
}

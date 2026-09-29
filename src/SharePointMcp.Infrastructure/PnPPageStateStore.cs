using System.Text.Json;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

public sealed record PnPPageApplicationState(
    string SiteIdentifier,
    string PageName,
    string DefinitionHash,
    string LivePageHash,
    DateTimeOffset AppliedAt);

public interface IPnPPageStateStore
{
    Task<PnPPageApplicationState?> ReadAsync(string siteIdentifier, string pageName, CancellationToken cancellationToken = default);
    Task WriteAsync(PnPPageApplicationState state, CancellationToken cancellationToken = default);
}

public sealed class JsonPnPPageStateStore(string path) : IPnPPageStateStore
{
    private readonly SemaphoreSlim _gate = new(1, 1);

    public async Task<PnPPageApplicationState?> ReadAsync(string siteIdentifier, string pageName, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var states = await ReadAllAsync(cancellationToken).ConfigureAwait(false);
            return states.FirstOrDefault(state =>
                string.Equals(state.SiteIdentifier, siteIdentifier, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(state.PageName, pageName, StringComparison.OrdinalIgnoreCase));
        }
        finally { _gate.Release(); }
    }

    public async Task WriteAsync(PnPPageApplicationState state, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var states = await ReadAllAsync(cancellationToken).ConfigureAwait(false);
            states.RemoveAll(value =>
                string.Equals(value.SiteIdentifier, state.SiteIdentifier, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(value.PageName, state.PageName, StringComparison.OrdinalIgnoreCase));
            states.Add(state);
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            var temporary = path + ".tmp";
            await File.WriteAllTextAsync(temporary, JsonSerializer.Serialize(states, JsonDefaults.Options), cancellationToken).ConfigureAwait(false);
            File.Move(temporary, path, true);
        }
        finally { _gate.Release(); }
    }

    private async Task<List<PnPPageApplicationState>> ReadAllAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(path)) return [];
        await using var input = File.OpenRead(path);
        return await JsonSerializer.DeserializeAsync<List<PnPPageApplicationState>>(input, JsonDefaults.Options, cancellationToken).ConfigureAwait(false) ?? [];
    }
}

internal sealed class NullPnPPageStateStore : IPnPPageStateStore
{
    public Task<PnPPageApplicationState?> ReadAsync(string siteIdentifier, string pageName, CancellationToken cancellationToken = default) => Task.FromResult<PnPPageApplicationState?>(null);
    public Task WriteAsync(PnPPageApplicationState state, CancellationToken cancellationToken = default) => Task.CompletedTask;
}

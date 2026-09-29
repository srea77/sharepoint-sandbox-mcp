using System.Text.Json;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

public sealed record ThemeApplicationState(
    string SiteIdentifier,
    SiteThemeDefinition Theme,
    string DefinitionHash,
    string LiveThemeInfoHash,
    DateTimeOffset AppliedAt);

public interface IThemeStateStore
{
    Task<ThemeApplicationState?> ReadAsync(string siteIdentifier, CancellationToken cancellationToken = default);
    Task WriteAsync(ThemeApplicationState state, CancellationToken cancellationToken = default);
}

public sealed class JsonThemeStateStore(string path) : IThemeStateStore
{
    private readonly SemaphoreSlim _gate = new(1, 1);

    public async Task<ThemeApplicationState?> ReadAsync(string siteIdentifier, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (!File.Exists(path)) return null;
            await using var input = File.OpenRead(path);
            var states = await JsonSerializer.DeserializeAsync<List<ThemeApplicationState>>(input, JsonDefaults.Options, cancellationToken).ConfigureAwait(false) ?? [];
            return states.FirstOrDefault(state => string.Equals(state.SiteIdentifier, siteIdentifier, StringComparison.OrdinalIgnoreCase));
        }
        finally { _gate.Release(); }
    }

    public async Task WriteAsync(ThemeApplicationState state, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            List<ThemeApplicationState> states = [];
            if (File.Exists(path))
            {
                await using var input = File.OpenRead(path);
                states = await JsonSerializer.DeserializeAsync<List<ThemeApplicationState>>(input, JsonDefaults.Options, cancellationToken).ConfigureAwait(false) ?? [];
            }
            states.RemoveAll(value => string.Equals(value.SiteIdentifier, state.SiteIdentifier, StringComparison.OrdinalIgnoreCase));
            states.Add(state);
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            var temporary = path + ".tmp";
            await File.WriteAllTextAsync(temporary, JsonSerializer.Serialize(states, JsonDefaults.Options), cancellationToken).ConfigureAwait(false);
            File.Move(temporary, path, true);
        }
        finally { _gate.Release(); }
    }
}

internal sealed class NullThemeStateStore : IThemeStateStore
{
    public Task<ThemeApplicationState?> ReadAsync(string siteIdentifier, CancellationToken cancellationToken = default) => Task.FromResult<ThemeApplicationState?>(null);
    public Task WriteAsync(ThemeApplicationState state, CancellationToken cancellationToken = default) => Task.CompletedTask;
}

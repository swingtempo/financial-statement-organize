namespace StatementOrganizer;

/// <summary>
/// Minimal .env loader. Existing environment variables always win.
/// </summary>
public static class EnvFile
{
    public static Dictionary<string, string> Load(params string[] candidatePaths)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in candidatePaths)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) continue;

            foreach (var raw in File.ReadAllLines(path))
            {
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith('#')) continue;

                var eq = line.IndexOf('=');
                if (eq <= 0) continue;

                var key = line[..eq].Trim();
                var value = line[(eq + 1)..].Trim();
                // strip optional quotes
                if (value.Length >= 2 &&
                    ((value[0] == '"' && value[^1] == '"') ||
                     (value[0] == '\'' && value[^1] == '\'')))
                {
                    value = value[1..^1];
                }
                // strip inline comments (only when unquoted and preceded by whitespace)
                var hashIdx = value.IndexOf(" #", StringComparison.Ordinal);
                if (hashIdx >= 0) value = value[..hashIdx].Trim();

                if (!values.ContainsKey(key)) values[key] = value;
            }
        }
        return values;
    }

    public static string Get(Dictionary<string, string> values, string key, string? defaultValue = null)
        => Environment.GetEnvironmentVariable(key)
           ?? (values.TryGetValue(key, out var v) && !string.IsNullOrWhiteSpace(v) ? v : defaultValue!)!;
}

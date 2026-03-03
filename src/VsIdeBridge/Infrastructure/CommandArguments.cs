using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace VsIdeBridge.Infrastructure;

internal sealed class CommandArguments
{
    private readonly Dictionary<string, List<string>> _values;

    public CommandArguments(Dictionary<string, List<string>> values)
    {
        _values = values;
    }

    public string? GetString(string name, string? defaultValue = null)
    {
        return _values.TryGetValue(name, out var values) && values.Count > 0
            ? values[values.Count - 1]
            : defaultValue;
    }

    public bool Has(string name)
    {
        return _values.ContainsKey(name);
    }

    public string GetRequiredString(string name)
    {
        var value = GetString(name);
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new CommandErrorException("invalid_arguments", $"Missing required argument --{name}.");
        }

        return value!;
    }

    public int GetInt32(string name, int defaultValue)
    {
        var raw = GetString(name);
        if (string.IsNullOrWhiteSpace(raw))
        {
            return defaultValue;
        }

        if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
        {
            throw new CommandErrorException("invalid_arguments", $"Argument --{name} must be an integer.");
        }

        return value;
    }

    public int? GetNullableInt32(string name)
    {
        var raw = GetString(name);
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
        {
            throw new CommandErrorException("invalid_arguments", $"Argument --{name} must be an integer.");
        }

        return value;
    }

    public bool GetBoolean(string name, bool defaultValue)
    {
        var raw = GetString(name);
        if (string.IsNullOrWhiteSpace(raw))
        {
            return defaultValue;
        }

        if (bool.TryParse(raw, out var value))
        {
            return value;
        }

        throw new CommandErrorException("invalid_arguments", $"Argument --{name} must be true or false.");
    }

    public string GetEnum(string name, string defaultValue, params string[] allowedValues)
    {
        var value = GetString(name, defaultValue) ?? defaultValue;
        if (!allowedValues.Contains(value, StringComparer.OrdinalIgnoreCase))
        {
            throw new CommandErrorException("invalid_arguments", $"Argument --{name} must be one of: {string.Join(", ", allowedValues)}.");
        }

        return value.ToLowerInvariant();
    }

    public IReadOnlyDictionary<string, List<string>> RawValues => _values;
}

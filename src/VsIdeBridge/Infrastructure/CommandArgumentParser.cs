using System.Collections.Generic;
using System.Text;

namespace VsIdeBridge.Infrastructure;

internal static class CommandArgumentParser
{
    public static CommandArguments Parse(string? rawArguments)
    {
        var tokens = Tokenize(rawArguments ?? string.Empty);
        var values = new Dictionary<string, List<string>>(System.StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];
            if (!token.StartsWith("--", System.StringComparison.Ordinal))
            {
                throw new CommandErrorException("invalid_arguments", $"Unexpected token '{token}'. Arguments must use --name value form.");
            }

            var name = token.Substring(2);
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new CommandErrorException("invalid_arguments", "Encountered an empty argument name.");
            }

            string value;
            if (i + 1 < tokens.Count && !tokens[i + 1].StartsWith("--", System.StringComparison.Ordinal))
            {
                value = tokens[++i];
            }
            else
            {
                value = "true";
            }

            if (!values.TryGetValue(name, out var list))
            {
                list = new List<string>();
                values[name] = list;
            }

            list.Add(value);
        }

        return new CommandArguments(values);
    }

    private static List<string> Tokenize(string text)
    {
        var tokens = new List<string>();
        var buffer = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            if (ch == '\\' && i + 1 < text.Length)
            {
                var next = text[i + 1];
                if (next == '"' || next == '\\')
                {
                    buffer.Append(next);
                    i++;
                    continue;
                }
            }

            if (ch == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (!inQuotes && char.IsWhiteSpace(ch))
            {
                FlushToken(tokens, buffer);
                continue;
            }

            buffer.Append(ch);
        }

        if (inQuotes)
        {
            throw new CommandErrorException("invalid_arguments", "Unterminated quoted argument.");
        }

        FlushToken(tokens, buffer);
        return tokens;
    }

    private static void FlushToken(List<string> tokens, StringBuilder buffer)
    {
        if (buffer.Length == 0)
        {
            return;
        }

        tokens.Add(buffer.ToString());
        buffer.Clear();
    }
}

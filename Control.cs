// unset

using System.Collections.Generic;
using System.Reflection.Metadata.Ecma335;

using NPOI.POIFS.Properties;

namespace shtxt
{
    public enum Compairator
    {
        Less,
        LessOrEqual,
        Greater,
        GreaterOrEqual,
        Equal,
        NotEqual,
    }
    public abstract record Control;

    public sealed record None : Control;
    public sealed record Comment : Control;

    public sealed record Version(Compairator compairator, string version) : Control;

    public class ControlParser
    {
        static IReadOnlyDictionary<string, Compairator> compairators = new Dictionary<string, Compairator>()
        {
            {"<=|", Compairator.LessOrEqual},
            {"<|", Compairator.Less},
            {">=|", Compairator.GreaterOrEqual},
            {">|", Compairator.Greater},
            {"=|", Compairator.Equal},
            {"!=|", Compairator.NotEqual},
        };

        public string CommentStartsWith { get; set; } = "";
        public Control Parse(string s)
        {
            if (s.StartsWith(CommentStartsWith))
            {
                return new Comment();
            }

            foreach (var elm in compairators)
            {
                if (s.StartsWith(elm.Key))
                {
                    return new Version(elm.Value, s.Substring(elm.Key.Length));
                }
            }

            return new None();
        }
    }
}
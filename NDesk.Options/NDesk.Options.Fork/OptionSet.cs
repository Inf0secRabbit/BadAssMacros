namespace NDesk.Options.Fork
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Linq;
    using Common;
    using ActionOptions;

    public class OptionSet : KeyedCollection<string, Option>
    {
        private const int OptionWidth = 29;

        private readonly Regex valueOption = new Regex(@"^(?<flag>--|-|/)(?<name>[^:=]+)((?<sep>[:=])(?<value>.*))?$");

        public OptionSet()
            : this(f => f)
        {
        }

        public OptionSet(Converter<string, string> localizer)
        {
            MessageLocalizer = localizer;
        }

        public Converter<string, string> MessageLocalizer { get; set; }

        protected override string GetKeyForItem(Option item)
        {
            if (item == null)
            {
                throw new ArgumentNullException("item");
            }

            if (item.Names != null && item.Names.Length > 0)
            {
                return item.Names[0];
            }

            // This should never happen, as it's invalid for Option to be
            // constructed w/o any names.
            throw new InvalidOperationException("Option has no names!");
        }

        protected override void InsertItem(int index, Option item)
        {
            base.InsertItem(index, item);
            AddImpl(item);
        }

        protected override void RemoveItem(int index)
        {
            base.RemoveItem(index);
            var p = Items[index];

            // KeyedCollection.RemoveItem() handles the 0th item
            for (var i = 1; i < p.Names.Length; ++i)
            {
                Dictionary.Remove(p.Names[i]);
            }
        }

        protected override void SetItem(int index, Option item)
        {
            base.SetItem(index, item);
            RemoveItem(index);
            AddImpl(item);
        }

        private void AddImpl(Option option)
        {
            if (option == null)
            {
                throw new ArgumentNullException("option");
            }

            var added = new List<string>(option.Names.Length);
            try
            {
                // KeyedCollection.InsertItem/SetItem handle the 0th name.
                for (var i = 1; i < option.Names.Length; ++i)
                {
                    Dictionary.Add(option.Names[i], option);
                    added.Add(option.Names[i]);
                }
            }
            catch (Exception)
            {
                foreach (var name in added)
                {
                    Dictionary.Remove(name);
                }

                throw;
            }
        }

        public new OptionSet Add(Option option)
        {
            base.Add(option);
            return this;
        }

        public OptionSet Add(string prototype, Action<string> action)
        {
            return Add(prototype, null, action);
        }

        public OptionSet Add(string prototype, string description, Action<string> action)
        {
            if (action == null)
            {
                throw new ArgumentNullException("action");
            }

            Option p = new ActionOption(prototype, description, 1, delegate (OptionValueCollection v) { action(v[0]); });
            base.Add(p);
            return this;
        }

        public OptionSet Add(string prototype, OptionAction<string, string> action)
        {
            return Add(prototype, null, action);
        }

        public OptionSet Add(string prototype, string description, OptionAction<string, string> action)
        {
            if (action == null)
            {
                throw new ArgumentNullException("action");
            }

            Option p = new ActionOption(
                prototype,
                description,
                2,
                delegate (OptionValueCollection v) { action(v[0], v[1]); });
            base.Add(p);
            return this;
        }

        public OptionSet Add<T>(string prototype, Action<T> action)
        {
            return Add(prototype, null, action);
        }

        public OptionSet Add<T>(string prototype, string description, Action<T> action)
        {
            return Add(new ActionOption<T>(prototype, description, action));
        }

        public OptionSet Add<TKey, TValue>(string prototype, OptionAction<TKey, TValue> action)
        {
            return Add(prototype, null, action);
        }

        public OptionSet Add<TKey, TValue>(string prototype, string description, OptionAction<TKey, TValue> action)
        {
            return Add(new ActionOption<TKey, TValue>(prototype, description, action));
        }

        protected virtual OptionContext CreateOptionContext()
        {
            return new OptionContext(this);
        }

        public List<string> Parse(IEnumerable<string> arguments)
        {
            var c = CreateOptionContext();
            c.OptionIndex = -1;
            var process = true;
            var unprocessed = new List<string>();
            var def = Contains("<>") ? this["<>"] : null;
            foreach (var argument in arguments)
            {
                ++c.OptionIndex;
                if (argument == "--")
                {
                    process = false;
                    continue;
                }

                if (!process)
                {
                    Unprocessed(unprocessed, def, c, argument);
                    continue;
                }

                if (!Parse(argument, c))
                {
                    // is it a bundled option?
                    var unprocessedBundleOptions = ParceBundle(argument, c);

                    if (unprocessedBundleOptions == null)
                    {
                        Unprocessed(unprocessed, def, c, argument);
                    }
                    else
                    {
                        unprocessedBundleOptions.ForEach(ar => Unprocessed(unprocessed, def, c, ar));
                    }
                }
            }

            if (c.Option != null)
            {
                c.Option.Invoke(c);
            }

            return unprocessed;
        }

        private static void Unprocessed(ICollection<string> extra, Option def, OptionContext c, string argument)
        {
            if (def == null)
            {
                extra.Add(argument);
                return;
            }

            c.OptionValues.Add(argument);
            c.Option = def;
            c.Option.Invoke(c);
        }

        protected bool GetOptionParts(string argument, out string flag, out string name, out string sep, out string value)
        {
            if (argument == null)
            {
                throw new ArgumentNullException("argument");
            }

            flag = name = sep = value = null;
            var m = valueOption.Match(argument);
            if (!m.Success)
            {
                return false;
            }

            flag = m.Groups["flag"].Value;
            name = m.Groups["name"].Value;
            if (m.Groups["sep"].Success && m.Groups["value"].Success)
            {
                sep = m.Groups["sep"].Value;
                value = m.Groups["value"].Value;
            }

            return true;
        }

        protected virtual bool Parse(string argument, OptionContext c)
        {
            if (c.Option != null)
            {
                ParseValue(argument, c);
                return true;
            }

            string f, n, s, v;
            if (!GetOptionParts(argument, out f, out n, out s, out v))
            {
                return false;
            }

            if (Contains(n))
            {
                var p = this[n];
                c.OptionName = f + n;
                c.Option = p;
                switch (p.OptionValueType)
                {
                    case OptionValueType.None:
                        c.OptionValues.Add(n);
                        c.Option.Invoke(c);
                        break;

                    case OptionValueType.Optional:
                    case OptionValueType.Required:
                        ParseValue(v, c);
                        break;

                    default:
                        throw new ArgumentOutOfRangeException("argument");
                }

                return true;
            }

            // no match; is it a bool option?
            // ReSharper disable once ConvertIfStatementToReturnStatement
            return ParseBool(argument, n, c);
        }

        private List<string> ParceBundle(string argument, OptionContext c)
        {
            string f, n, s, v;
            var result = new List<string>();

            if (!GetOptionParts(argument, out f, out n, out s, out v))
            {
                return null;
            }

            if (f != "-")
            {
                return null;
            }

            for (var i = 0; i < n.Length; ++i)
            {
                var opt = f + n[i];
                var rn = n[i].ToString();
                if (!Contains(rn))
                {
                    result.Add(opt);
                    continue;
                }

                var p = this[rn];
                switch (p.OptionValueType)
                {
                    case OptionValueType.None:
                        Invoke(c, opt, n, p);
                        break;

                    case OptionValueType.Optional:
                    case OptionValueType.Required:
                        {
                            var option = n.Substring(i + 1);
                            c.Option = p;
                            c.OptionName = opt;
                            ParseValue(option.Length != 0 ? option : null, c);
                        }
                        break;

                    default:
                        throw new InvalidOperationException("Unknown OptionValueType: " + p.OptionValueType);
                }
            }
            return result;
        }

        private void ParseValue(string option, OptionContext c)
        {
            if (option != null)
            {
                foreach (var o in
                    c.Option.ValueSeparators != null
                        ? option.Split(c.Option.ValueSeparators, StringSplitOptions.None)
                        : new[] { option })
                {
                    c.OptionValues.Add(o);
                }
            }

            if (c.OptionValues.Count == c.Option.MaxValueCount || c.Option.OptionValueType == OptionValueType.Optional)
            {
                c.Option.Invoke(c);
            }
            else if (c.OptionValues.Count > c.Option.MaxValueCount)
            {
                throw new OptionException(
                    MessageLocalizer(
                        string.Format(
                            "Error: Found {0} option values when expecting {1}.",
                            c.OptionValues.Count,
                            c.Option.MaxValueCount)),
                    c.OptionName);
            }
        }

        private bool ParseBool(string option, string n, OptionContext c)
        {
            string rn;
            if (n.Length < 1 || (n[n.Length - 1] != '+' && n[n.Length - 1] != '-')
                || !Contains(rn = n.Substring(0, n.Length - 1)))
            {
                return false;
            }

            var p = this[rn];
            var v = n[n.Length - 1] == '+' ? option : null;
            c.OptionName = option;
            c.Option = p;
            c.OptionValues.Add(v);
            p.Invoke(c);
            return true;
        }

        private static void Invoke(OptionContext c, string name, string value, Option option)
        {
            c.OptionName = name;
            c.Option = option;
            c.OptionValues.Add(value);
            option.Invoke(c);
        }

        public void WriteOptionDescriptions(TextWriter o)
        {
            foreach (var p in this)
            {
                var written = 0;
                if (!WriteOptionPrototype(o, p, ref written))
                {
                    continue;
                }

                if (written < OptionWidth)
                {
                    o.Write(new string(' ', OptionWidth - written));
                }
                else
                {
                    o.WriteLine();
                    o.Write(new string(' ', OptionWidth));
                }

                var lines = GetLines(MessageLocalizer(GetDescription(p.Description)));
                o.WriteLine(lines[0]);
                var prefix = new string(' ', OptionWidth + 2);
                for (var i = 1; i < lines.Count; ++i)
                {
                    o.Write(prefix);
                    o.WriteLine(lines[i]);
                }
            }
        }

        private bool WriteOptionPrototype(TextWriter o, Option p, ref int written)
        {
            var names = p.Names;

            var i = GetNextOptionIndex(names, 0);
            if (i == names.Length)
            {
                return false;
            }

            if (names[i].Length == 1)
            {
                Write(o, ref written, "  -");
                Write(o, ref written, names[0]);
            }
            else
            {
                Write(o, ref written, "      --");
                Write(o, ref written, names[0]);
            }

            for (i = GetNextOptionIndex(names, i + 1); i < names.Length; i = GetNextOptionIndex(names, i + 1))
            {
                Write(o, ref written, ", ");
                Write(o, ref written, names[i].Length == 1 ? "-" : "--");
                Write(o, ref written, names[i]);
            }

            if (p.OptionValueType == OptionValueType.Optional || p.OptionValueType == OptionValueType.Required)
            {
                if (p.OptionValueType == OptionValueType.Optional)
                {
                    Write(o, ref written, MessageLocalizer("["));
                }

                Write(o, ref written, MessageLocalizer("=" + GetArgumentName(0, p.MaxValueCount, p.Description)));
                var sep = p.ValueSeparators != null && p.ValueSeparators.Length > 0 ? p.ValueSeparators[0] : " ";
                for (var c = 1; c < p.MaxValueCount; ++c)
                {
                    Write(
                        o,
                        ref written,
                        MessageLocalizer(sep + GetArgumentName(c, p.MaxValueCount, p.Description)));
                }

                if (p.OptionValueType == OptionValueType.Optional)
                {
                    Write(o, ref written, MessageLocalizer("]"));
                }
            }

            return true;
        }

        private static int GetNextOptionIndex(string[] names, int i)
        {
            while (i < names.Length && names[i] == "<>")
            {
                ++i;
            }

            return i;
        }

        private static void Write(TextWriter o, ref int n, string s)
        {
            n += s.Length;
            o.Write(s);
        }

        private static string GetArgumentName(int index, int maxIndex, string description)
        {
            if (description == null)
            {
                return maxIndex == 1 ? "VALUE" : "VALUE" + (index + 1);
            }

            var nameStart = maxIndex == 1 ? new[] { "{0:", "{" } : new[] { "{" + index + ":" };

            foreach (var str in nameStart)
            {
                int start, j = 0;
                do
                {
                    start = description.IndexOf(str, j, StringComparison.Ordinal);
                }
                while (start >= 0 && j != 0 && description[j++ - 1] == '{');
                if (start == -1)
                {
                    continue;
                }

                var end = description.IndexOf("}", start, StringComparison.Ordinal);
                if (end == -1)
                {
                    continue;
                }

                return description.Substring(start + str.Length, end - start - str.Length);
            }

            return maxIndex == 1 ? "VALUE" : "VALUE" + (index + 1);
        }

        private static string GetDescription(string description)
        {
            if (description == null)
            {
                return string.Empty;
            }

            var sb = new StringBuilder(description.Length);
            var start = -1;
            for (var i = 0; i < description.Length; ++i)
            {
                switch (description[i])
                {
                    case '{':
                        if (i == start)
                        {
                            sb.Append('{');
                            start = -1;
                        }
                        else if (start < 0)
                        {
                            start = i + 1;
                        }

                        break;

                    case '}':
                        if (start < 0)
                        {
                            if (i + 1 == description.Length || description[i + 1] != '}')
                            {
                                throw new InvalidOperationException("Invalid option description: " + description);
                            }

                            ++i;
                            sb.Append("}");
                        }
                        else
                        {
                            sb.Append(description.Substring(start, i - start));
                            start = -1;
                        }

                        break;

                    case ':':
                        if (start < 0)
                        {
                            goto default;
                        }

                        start = i + 1;
                        break;

                    default:
                        if (start < 0)
                        {
                            sb.Append(description[i]);
                        }

                        break;
                }
            }

            return sb.ToString();
        }

        private static List<string> GetLines(string description)
        {
            var lines = new List<string>();
            if (string.IsNullOrEmpty(description))
            {
                lines.Add(string.Empty);
                return lines;
            }

            var length = 80 - OptionWidth - 2;
            int start = 0, end;
            do
            {
                end = GetLineEnd(start, length, description);
                var cont = false;
                if (end < description.Length)
                {
                    var c = description[end];
                    if (c == '-' || (char.IsWhiteSpace(c) && c != '\n'))
                    {
                        ++end;
                    }
                    else if (c != '\n')
                    {
                        cont = true;
                        --end;
                    }
                }

                lines.Add(description.Substring(start, end - start));
                if (cont)
                {
                    lines[lines.Count - 1] += "-";
                }

                start = end;
                if (start < description.Length && description[start] == '\n')
                {
                    ++start;
                }
            }
            while (end < description.Length);
            return lines;
        }

        private static int GetLineEnd(int start, int length, string description)
        {
            var end = Math.Min(start + length, description.Length);
            var sep = -1;
            for (var i = start; i < end; ++i)
            {
                switch (description[i])
                {
                    case ' ':
                    case '\t':
                    case '\v':
                    case '-':
                    case ',':
                    case '.':
                    case ';':
                        sep = i;
                        break;

                    case '\n':
                        return i;
                }
            }

            if (sep == -1 || end == description.Length)
            {
                return end;
            }

            return sep;
        }
    }
}
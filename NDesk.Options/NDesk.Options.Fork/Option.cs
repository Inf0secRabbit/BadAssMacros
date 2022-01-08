namespace NDesk.Options.Fork
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using Common;

    public abstract class Option
    {
        private static readonly char[] NameTerminator = { '=', ':' };

        protected Option(string prototype, string description)
            : this(prototype, description, 1)
        {
        }

        protected Option(string prototype, string description, int maxValueCount)
        {
            if (prototype == null)
            {
                throw new ArgumentNullException("prototype");
            }

            if (prototype.Length == 0)
            {
                throw new ArgumentException("Cannot be the empty string.", "prototype");
            }

            if (maxValueCount < 0)
            {
                throw new ArgumentOutOfRangeException("maxValueCount");
            }

            Prototype = prototype;
            Names = prototype.Split('|');
            Description = description;
            MaxValueCount = maxValueCount;
            OptionValueType = ParsePrototype();

            if (MaxValueCount == 0 && OptionValueType != OptionValueType.None)
            {
                throw new ArgumentException(
                    "Cannot provide maxValueCount of 0 for OptionValueType.Required or " + "OptionValueType.Optional.",
                    "maxValueCount");
            }

            if (OptionValueType == OptionValueType.None && maxValueCount > 1)
            {
                throw new ArgumentException(
                    string.Format("Cannot provide maxValueCount of {0} for OptionValueType.None.", maxValueCount),
                    "maxValueCount");
            }

            if (Array.IndexOf(Names, "<>") >= 0
                && ((Names.Length == 1 && OptionValueType != OptionValueType.None)
                    || (Names.Length > 1 && MaxValueCount > 1)))
            {
                throw new ArgumentException("The default option handler '<>' cannot require values.", "prototype");
            }
        }

        public string Description { get; set; }

        public int MaxValueCount { get; set; }

        public OptionValueType OptionValueType { get; set; }

        public string Prototype { get; set; }

        internal string[] Names { get; set; }

        internal string[] ValueSeparators { get; private set; }

        public string[] GetNames()
        {
            return (string[])Names.Clone();
        }

        public string[] GetValueSeparators()
        {
            if (ValueSeparators == null)
            {
                return new string[0];
            }

            return (string[])ValueSeparators.Clone();
        }

        public void Invoke(OptionContext c)
        {
            OnParseComplete(c);
            c.OptionName = null;
            c.Option = null;
            c.OptionValues.Clear();
        }

        public override string ToString()
        {
            return Prototype;
        }

        protected static T Parse<T>(string value, OptionContext c)
        {
            var conv = TypeDescriptor.GetConverter(typeof(T));
            var t = default(T);
            try
            {
                if (value != null)
                {
                    t = (T)conv.ConvertFromString(value);
                }
            }
            catch (Exception e)
            {
                throw new OptionException(
                    string.Format(
                        c.OptionSet.MessageLocalizer("Could not convert string `{0}' to type {1} for option `{2}'."),
                        value,
                        typeof(T).Name,
                        c.OptionName),
                    c.OptionName,
                    e);
            }

            return t;
        }

        protected abstract void OnParseComplete(OptionContext c);

        private static void AddSeparators(string name, int end, ICollection<string> seps)
        {
            var start = -1;
            for (var i = end + 1; i < name.Length; ++i)
            {
                switch (name[i])
                {
                    case '{':
                        if (start != -1)
                        {
                            throw new ArgumentException(
                                string.Format("Ill-formed name/value separator found in \"{0}\".", name),
                                "name");
                        }

                        start = i + 1;
                        break;

                    case '}':
                        if (start == -1)
                        {
                            throw new ArgumentException(
                                string.Format("Ill-formed name/value separator found in \"{0}\".", name),
                                "name");
                        }

                        seps.Add(name.Substring(start, i - start));
                        start = -1;
                        break;

                    default:
                        if (start == -1)
                        {
                            seps.Add(name[i].ToString());
                        }

                        break;
                }
            }

            if (start != -1)
            {
                throw new ArgumentException(
                    string.Format("Ill-formed name/value separator found in \"{0}\".", name));
            }
        }

        private OptionValueType ParsePrototype()
        {
            var type = '\0';
            var seps = new List<string>();
            for (var i = 0; i < Names.Length; ++i)
            {
                var name = Names[i];
                if (name.Length == 0)
                {
                    throw new ArgumentException("Empty option names are not supported.");
                }

                var end = name.IndexOfAny(NameTerminator);
                if (end == -1)
                {
                    continue;
                }

                Names[i] = name.Substring(0, end);
                if (type == '\0' || type == name[end])
                {
                    type = name[end];
                }
                else
                {
                    throw new ArgumentException(
                        string.Format("Conflicting option types: '{0}' vs. '{1}'.", type, name[end]));
                }

                AddSeparators(name, end, seps);
            }

            if (type == '\0')
            {
                return OptionValueType.None;
            }

            if (MaxValueCount <= 1 && seps.Count != 0)
            {
                throw new ArgumentException(
                    string.Format(
                        "Cannot provide key/value separators for Options taking {0} value(s).",
                        MaxValueCount));
            }

            if (MaxValueCount > 1)
            {
                if (seps.Count == 0)
                {
                    ValueSeparators = new[] { ":", "=" };
                }
                else if (seps.Count == 1 && seps[0].Length == 0)
                {
                    ValueSeparators = null;
                }
                else
                {
                    ValueSeparators = seps.ToArray();
                }
            }

            return type == '=' ? OptionValueType.Required : OptionValueType.Optional;
        }
    }
}
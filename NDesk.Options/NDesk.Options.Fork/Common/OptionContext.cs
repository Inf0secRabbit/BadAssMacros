namespace NDesk.Options.Fork.Common
{
    public class OptionContext
    {
        public OptionContext(OptionSet set)
        {
            OptionSet = set;
            OptionValues = new OptionValueCollection(this);
        }

        public Option Option { get; set; }

        public int OptionIndex { get; set; }

        public string OptionName { get; set; }

        public OptionSet OptionSet { get; set; }

        public OptionValueCollection OptionValues { get; set; }
    }
}
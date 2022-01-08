namespace NDesk.Options.Fork.ActionOptions
{
    using System;
    using Common;

    internal class ActionOption<T> : Option
    {
        private readonly Action<T> action;

        public ActionOption(string prototype, string description, Action<T> action)
            : base(prototype, description, 1)
        {
            if (action == null)
            {
                throw new ArgumentNullException("action");
            }

            this.action = action;
        }

        protected override void OnParseComplete(OptionContext c)
        {
            action(Parse<T>(c.OptionValues[0], c));
        }
    }
}
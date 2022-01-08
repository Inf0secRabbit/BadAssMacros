namespace NDesk.Options.Fork.ActionOptions
{
    using System;
    using Common;

    internal class ActionOption<TKey, TValue> : Option
    {
        private readonly OptionAction<TKey, TValue> action;

        public ActionOption(string prototype, string description, OptionAction<TKey, TValue> action)
            : base(prototype, description, 2)
        {
            if (action == null)
            {
                throw new ArgumentNullException("action");
            }

            this.action = action;
        }

        protected override void OnParseComplete(OptionContext c)
        {
            action(Parse<TKey>(c.OptionValues[0], c), Parse<TValue>(c.OptionValues[1], c));
        }
    }
}
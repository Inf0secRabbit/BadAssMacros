namespace NDesk.Options.Fork.Common
{
    public delegate void OptionAction<in TKey, in TValue>(TKey key, TValue value);
}
namespace System.Windows
{
    internal class Forms
    {
        public static object DialogResult { get; internal set; }

        internal class ColorDialog
        {
            internal object ShowDialog()
            {
                throw new NotImplementedException();
            }
        }
    }
}
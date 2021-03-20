using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxGen
{
    public abstract class Replacement
    {
        public abstract override string ToString();
    }

    public class SimpleReplacement : Replacement
    {
        private string text;

        public SimpleReplacement(string text)
        {
            this.text = text;
        }

        public override string ToString()
        {
            return text;
        }
    }

    public class DateTimeReplacement : Replacement
    {
        private string format;

        public DateTimeReplacement(string format = "MM/dd/yyyy")
        {
            this.format = format;
        }

        public override string ToString()
        {
            return DateTime.Now.ToString(format);
        }
    }
}

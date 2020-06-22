using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DocJournalParser
{
    public class JDiscription
    {
        public Autor Autor { get; set; } = new Autor();
        public Editor FirstEditor { get; set; } = new Editor();
        public string Title { get; set; } = string.Empty;
        public string TitleInfo { get; set; } = string.Empty;
        public string Editors { get; set; } = string.Empty;
        public string Year { get; set; } = string.Empty;
        public string JVolume { get; set; } = string.Empty;
        public string JNumber { get; set; } = string.Empty;
        public string Pages { get; set; } = string.Empty;
        public string Pagination { get; set; } = string.Empty;
        
        public string Notes { get; set; } = string.Empty;
        public string FullPublication { get; set; } = string.Empty;
        public string FullPubYear { get; set; } = string.Empty;
        public string FullPubVolume { get; set; } = string.Empty;
        public string FullPubNumber { get; set; } = string.Empty;
        public string FullPubPageRange { get; set; } = string.Empty;
        public dynamic DеscriptionNumber { get; set; } = string.Empty;

        private PropertyInfo[] _PropertyInfos = null;

        public override string ToString()
        {
            if (_PropertyInfos == null)
                _PropertyInfos = GetType().GetProperties();

            var builder = new StringBuilder();

            foreach (var info in _PropertyInfos)
            {
                var value = info.GetValue(this, null) ?? "(null)";
                builder.AppendLine(info.Name + ": " + value.ToString());
            }

            return builder.ToString();
        }
    }
}
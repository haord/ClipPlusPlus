using System.Collections.Generic;

namespace ClipPlusPlus.Models
{
    public class HistoryGroup
    {
        public string Name { get; set; } = string.Empty;
        public List<HistoryItem> Items { get; set; } = new();
    }
}

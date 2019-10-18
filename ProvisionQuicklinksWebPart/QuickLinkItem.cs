using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProvisionQuicklinksWebPart
{
    public enum ThumbnailType
    {
        Image = 1,
        Icon = 2
    }
    public class QuickLinkItem
    {
        public string UniqueId { get; set; }
        public string Url { get; set; }
        public string Title { get; set; }
        public string Description { get; set; } = string.Empty;
        public string AltText { get; set; } = string.Empty;
        public string ImageUrl { get; set; }
        public ThumbnailType ThumbnailType { get; set; }
        public string IconName { get; set; }
    }
}

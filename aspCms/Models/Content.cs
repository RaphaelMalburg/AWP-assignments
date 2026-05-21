using System.ComponentModel.DataAnnotations;

namespace aspCms.Models
{
    public class Content
    {
        public int Id { get; set; }

        [Required]
        public string Page { get; set; } = string.Empty;

        [Required]
        public string Title { get; set; } = string.Empty;

        [Required]
        public string Text { get; set; } = string.Empty;

        public string Author { get; set; } = string.Empty;

        public DateTime Date { get; set; } = DateTime.Now;
    }
}
using Microsoft.EntityFrameworkCore;
using aspCms.Models;

namespace aspCms.Data
{
    public class DbMiniCMSContext : DbContext
    {
        public DbMiniCMSContext(DbContextOptions<DbMiniCMSContext> options)
            : base(options)
        {
        }

        public DbSet<Content> Contents { get; set; }
    }
}
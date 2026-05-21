using aspCms.Models;

namespace aspCms.Data
{
    public static class DbInitializer
    {
        public static void Initialize(DbMiniCMSContext context)
        {
            if (context.Contents.Any())
            {
                return;
            }

            var homeContent = new Content
            {
                Page = "Home",
                Title = "<h1>Welcome to MiniCMS</h1>",
                Text = @"
                    <div class=""welcome-content"">
                        <p>This is the <b>miniCMS</b> homepage. You can manage content through the admin panel.</p>
                        <img src=""/images/cms-illustration.svg"" alt=""CMS Illustration"" class=""img-fluid"" />
                    </div>",
                Author = "System",
                Date = DateTime.Now
            };

            context.Contents.Add(homeContent);
            context.SaveChanges();
        }
    }
}
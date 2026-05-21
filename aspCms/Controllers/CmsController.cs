using Microsoft.AspNetCore.Mvc;
using aspCms.Data;

namespace aspCms.Controllers
{
    public class CmsController : Controller
    {
        private readonly DbMiniCMSContext _context;

        public CmsController(DbMiniCMSContext context)
        {
            _context = context;
        }

        public IActionResult Page(string id)
        {
            if (string.IsNullOrEmpty(id))
            {
                return NotFound();
            }

            var content = _context.Contents.FirstOrDefault(c => c.Page == id);
            
            if (content == null)
            {
                return NotFound();
            }

            return View(content);
        }
    }
}
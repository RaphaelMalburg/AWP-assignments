using Microsoft.AspNetCore.Mvc;
using aspCms.Data;
using aspCms.Models;

namespace aspCms.Controllers
{
    public class AdminController : Controller
    {
        private readonly DbMiniCMSContext _context;

        public AdminController(DbMiniCMSContext context)
        {
            _context = context;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Index(string username, string password)
        {
            if (username == "admin" && password == "admin")
            {
                HttpContext.Session.SetString("Login", "true");
                return RedirectToAction("List");
            }

            ViewBag.Error = "Invalid credentials";
            return View();
        }

        private bool IsLoggedIn()
        {
            return HttpContext.Session.GetString("Login") == "true";
        }

        public IActionResult List()
        {
            if (!IsLoggedIn())
            {
                return RedirectToAction("Index");
            }

            var contents = _context.Contents.ToList();
            return View(contents);
        }

        public IActionResult Create()
        {
            if (!IsLoggedIn())
            {
                return RedirectToAction("Index");
            }

            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(Content content)
        {
            if (!IsLoggedIn())
            {
                return RedirectToAction("Index");
            }

            if (ModelState.IsValid)
            {
                content.Date = DateTime.Now;
                _context.Add(content);
                _context.SaveChanges();
                return RedirectToAction("List");
            }

            return View(content);
        }

        public IActionResult Edit(int id)
        {
            if (!IsLoggedIn())
            {
                return RedirectToAction("Index");
            }

            var content = _context.Contents.Find(id);
            if (content == null)
            {
                return NotFound();
            }

            return View(content);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(int id, Content content)
        {
            if (!IsLoggedIn())
            {
                return RedirectToAction("Index");
            }

            if (id != content.Id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                content.Date = DateTime.Now;
                _context.Update(content);
                _context.SaveChanges();
                return RedirectToAction("List");
            }

            return View(content);
        }

        public IActionResult Delete(int id)
        {
            if (!IsLoggedIn())
            {
                return RedirectToAction("Index");
            }

            var content = _context.Contents.Find(id);
            if (content == null)
            {
                return NotFound();
            }

            return View(content);
        }

        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteConfirmed(int id)
        {
            if (!IsLoggedIn())
            {
                return RedirectToAction("Index");
            }

            var content = _context.Contents.Find(id);
            if (content != null)
            {
                _context.Contents.Remove(content);
                _context.SaveChanges();
            }

            return RedirectToAction("List");
        }

        public IActionResult Logout()
        {
            HttpContext.Session.Remove("Login");
            return RedirectToAction("Index");
        }
    }
}
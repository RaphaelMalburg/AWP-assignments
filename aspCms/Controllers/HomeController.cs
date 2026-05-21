using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using aspCms.Models;
using aspCms.Data;

namespace aspCms.Controllers;

public class HomeController : Controller
{
    private readonly DbMiniCMSContext _context;

    public HomeController(DbMiniCMSContext context)
    {
        _context = context;
    }

    public IActionResult Index()
    {
        var content = _context.Contents.FirstOrDefault(c => c.Page == "Home");
        return View(content);
    }

    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
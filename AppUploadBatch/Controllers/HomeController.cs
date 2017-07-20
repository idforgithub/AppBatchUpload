using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;

namespace AppBatchUpload.Controllers
{
    public class HomeController : Controller
    {
        [Authorize]
        [AllowAnonymous]
        public ViewResult Index()
        {
            return View();
        }
    }
}

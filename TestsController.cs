using GRC.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GRC.Controllers
{
    public class TestsController : BaseController
    {
        private GRCModelContainer db = new GRCModelContainer();
        //
        // GET: /Tests/
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Photo()
        {
            return View();
        }


        public ActionResult SavePhoto()
        {

            return View();
        }
        public ActionResult SavePhoto(FormCollection form)
        {
            var photo = new Photo();
            return View();
        }
        
	}
}
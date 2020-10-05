using GRC.Models;
using GRC.Models.Security;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

namespace GRC.Controllers
{
    public class SecurityController : BaseController
    {
        private GRCModelContainer Context = new GRCModelContainer();
        // GET: Security
        public ActionResult Login(string returnUrl = "")
        {
            return RedirectToAction("Index", new { returnUrl = returnUrl });
        }
        public ActionResult Index(string returnUrl = "")
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(GRC.Models.Security.LoginViewModel model, string returnUrl = "")
        {
            if (ModelState.IsValid)
            {
                var user = Context.Users.Where(u =>
                    u.UserName == model.Username && 
                    u.Password == model.Password).FirstOrDefault();
                if (user != null)
                {
                    var roles = user.Roles.Select(m => m.RoleName).ToArray();

                    CustomPrincipalSerializeModel serializeModel = new CustomPrincipalSerializeModel();
                    serializeModel.UserId = user.UserId;
                    serializeModel.PersonnelId = user.Personne.PersonnelId;
                    serializeModel.NomComplet = user.Personne.NomComplet;
                    serializeModel.Matricule = user.Personne.Matricule;
                    serializeModel.roles = roles;

                    string userData = JsonConvert.SerializeObject(serializeModel);
                    FormsAuthenticationTicket authTicket = new FormsAuthenticationTicket(
                        1,
                        user.Personne.Matricule,
                        DateTime.Now,
                        DateTime.Now.AddMinutes(15),
                        false,
                        userData
                        );
                             
                    string encTicket = FormsAuthentication.Encrypt(authTicket);
                    HttpCookie faCookie = new HttpCookie(FormsAuthentication.FormsCookieName, encTicket);
                    Response.Cookies.Add(faCookie);

                    //if (roles.Contains("Admin"))
                    //{
                    //    return RedirectToAction("Index", "Admin");
                    //}
                    //else if (roles.Contains("User"))
                    //{
                    //    return RedirectToAction("Index", "User");
                    //}
                    //else
                    //{
                    //    return RedirectToAction("Index", "Home");
                    //}
                    if (this.Url.IsLocalUrl(returnUrl) && returnUrl.Length > 1 && returnUrl.StartsWith("/")
                    && !returnUrl.StartsWith("//") && !returnUrl.StartsWith("/\\"))
                    {
                        return this.Redirect(returnUrl);
                    }
                    return RedirectToAction("Index", "Home");
                }

                ModelState.AddModelError("", "Incorrect username and/or password");
            }

            return View(model);
        }

        [AllowAnonymous]
        public ActionResult LogOut()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Login", "Security", null);
        }

        public ActionResult Manage()
        {
            var CurrentUser = System.Web.HttpContext.Current.User;
            var User = Context.Users.Where(x => x.Personne.Matricule == CurrentUser.Identity.Name).FirstOrDefault();
            if (User == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ChangePasswordViewModel users = new ChangePasswordViewModel
            {
                User = User
            };
            if (users.User == null)
            {
                return HttpNotFound();
            }
            //ViewBag.PersonnelId = new SelectList(db.PersonnelSet, "PersonnelId", "Nom", users.User.PersonnelId);
            return View(users);
        }

        // POST: Users/Edit/5
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost, ActionName("Manage")]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(ChangePasswordViewModel users)
        {
            if (ModelState.IsValid)
            {
                var user = Context.Users.Find(users.User.UserId);
                if (users.OldPassword == user.Password)
                {
                    user.Password = users.NewPassword;
                    Context.Entry(user).State = EntityState.Modified;
                    Context.SaveChanges();
                    TempData["success"] = "<li>Mot de passe changé</li>";
                    return RedirectToAction("Manage");
                }
            }
            //ViewBag.PersonnelId = new SelectList(db.PersonnelSet, "PersonnelId", "Nom", users.User.PersonnelId);
            return View(users);
        }
    }
}
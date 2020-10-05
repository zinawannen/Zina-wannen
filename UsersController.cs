using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using GRC.Models;
using GRC.Models.Security;

namespace GRC.Controllers
{
    public class UsersController : Controller
    {
        private GRCModelContainer db = new GRCModelContainer();

        // GET: Users
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Index()
        {
            var users = db.Users.Include(u => u.Personne);
            return View(users.ToList());
        }

        // GET: Users/Details/5
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            var roles = db.Roles.ToList();
            ViewBag.Roles = roles;

            Users users = db.Users.Find(id);
            if (users == null)
            {
                return HttpNotFound();
            }
            return View(users);
        }

        // GET: Users/Create
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Create()
        {
            ViewBag.PersonnelId = new SelectList(db.PersonnelSet.ToList().Select(x=>new { x.PersonnelId, Nom = x.Matricule + " | "+ x.NomComplet }), "PersonnelId", "Nom");
            return View(new Users { Password = "initial", IsActive = true, DateCreation = DateTime.Now });
        }

        // POST: Users/Create
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "UserId,UserName,Password,PersonnelId,IsActive,DateCreation")] Users users)
        {
            if (ModelState.IsValid)
            {
                users.DateCreation = DateTime.Now;
                users.IsActive = true;
                db.Users.Add(users);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.PersonnelId = new SelectList(db.PersonnelSet, "PersonnelId", "Nom", users.PersonnelId);
            return View(users);
        }

        // GET: Users/Edit/5
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            var roles = db.Roles.ToList();
            ViewBag.Roles = roles;

            ChangePasswordViewModel users = new ChangePasswordViewModel
            {
                User = db.Users.Find(id)
            };
            if (users.User == null)
            {
                return HttpNotFound();
            }
            ViewBag.PersonnelId = new SelectList(db.PersonnelSet, "PersonnelId", "Nom", users.User.PersonnelId);
            return View(users);
        }

        // POST: Users/Edit/5
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Edit(ChangePasswordViewModel users)
        {
            //http://www.entityframeworktutorial.net/EntityFramework4.3/validate-entity-in-entity-framework.aspx
            if (ModelState.IsValid)
            {
                var user = db.Users.Find(users.User.UserId);
                if (users.OldPassword == user.Password)
                {
                    user.Password = users.NewPassword;
                    db.Entry(user).State = EntityState.Modified;
                    db.SaveChanges();
                    return RedirectToAction("Index");
                }
            }
            ViewBag.PersonnelId = new SelectList(db.PersonnelSet, "PersonnelId", "Nom", users.User.PersonnelId);
            return View(users);
        }

        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult EditRoles(int id)
        {
            Users users = db.Users.Find(id);

            var roles = db.Roles.ToList();
            ViewBag.Roles = roles;
            ViewBag.EditRoles = true;

            return View("Details", users);
        }

        [HttpPost, ActionName("EditRoles")]
        public ActionResult EditRoles(int id, int[] role,  FormCollection form)
        {
            Users users = db.Users.Find(id);
            //var roles = role.ToList();
            //var rolesUser = users.Roles.Select(x => x.RoleId).ToList();

            //var rToRemove = rolesUser.Except(roles);
            //var rToAdd = roles.Except(rolesUser);

            users.Roles.Clear();
            if (role != null)
            {
                if (role.Count() > 0)
                {
                    foreach (var r in role)
                    {
                        users.Roles.Add(db.Roles.Find(r));
                    }
                }
            }
            db.SaveChanges();

            return RedirectToAction("Details", new { id = id });
        }

        // GET: Users/Delete/5
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Users users = db.Users.Find(id);
            if (users == null)
            {
                return HttpNotFound();
            }
            return View(users);
        }

        // POST: Users/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult DeleteConfirmed(int id)
        {
            Users users = db.Users.Find(id);
            db.Users.Remove(users);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
    
}

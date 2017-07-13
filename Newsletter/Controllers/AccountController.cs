using Newsletter.Models;
using System.Configuration;
using System.DirectoryServices.AccountManagement;
using System.Web.Mvc;
using System.Linq;
using System.Web.Security;

namespace Newsletter.Controllers
{
    [Authorize]
    public class AccountController : Controller
    {
        //
        // GET: /Account/Login
        [AllowAnonymous]
        public ActionResult Login(string returnUrl)
        {
            ViewBag.ReturnUrl = returnUrl;
            return View();
        }

        //
        // POST: /Account/Login
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult Login(LoginViewModel model, string returnUrl)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, ConfigurationManager.AppSettings["ADIP"].ToString()))
            {
                bool isValid = pc.ValidateCredentials(model.UserName, model.Password);

                if (isValid)
                {
                    // check if this user is attached to the newsletter group
                    var group = GroupPrincipal.FindByIdentity(pc, ConfigurationManager.AppSettings["NewsletterGroup"].ToString());
                    var isInGroup = group.GetMembers(true).Where(p => p.UserPrincipalName.ToLowerInvariant() == model.UserName.ToLowerInvariant() + "@ieianchorpensions.net").Any();
                    
                    if (!isInGroup)
                    {
                        ModelState.AddModelError("", "Sorry, you are not a member of the newsletter group");
                        return View(model);
                    }

                    string name = "", userEmail = "";
                    var usr = UserPrincipal.FindByIdentity(pc, model.UserName);
                    if (usr != null)
                    {
                        name = usr.DisplayName;
                        userEmail = usr.EmailAddress;
                    }

                    FormsAuthentication.SetAuthCookie(model.UserName, false);

                    Session["displayName"] = name;


                    if (Url.IsLocalUrl(returnUrl) && returnUrl.Length > 1
                        && returnUrl.StartsWith("/") && !returnUrl.StartsWith("//")
                        && !returnUrl.StartsWith("/\\"))
                    {
                        return Redirect(returnUrl);
                    }
                    else
                    {
                        return RedirectToAction("Index", "Home");
                    }
                }
                else
                {
                    ModelState.AddModelError("", "The user name or password provided is incorrect");
                }
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult AccessDenied()
        {
            return View();
        }

        //
        // POST: /Account/LogOff
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult LogOff()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Login", "Account");
        }
    }
}
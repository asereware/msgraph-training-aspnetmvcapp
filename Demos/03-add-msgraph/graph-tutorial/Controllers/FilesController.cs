using graph_tutorial.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace graph_tutorial.Controllers
{
    public class FilesController : BaseController
    {
        // GET: Files
        [Authorize]
        public async Task<ActionResult> Index()
        {
            var files = await GraphHelper.GetFilesAsync();

            // Change start and end dates from UTC to local time
            foreach (var fi in files)
            {
                fi.LastModifiedDateTime = fi.LastModifiedDateTime.HasValue 
                    ? fi.LastModifiedDateTime.Value.ToLocalTime()
                    : (DateTimeOffset?)null;

            }

            return View(files);
        }

        [Authorize]
        public async Task<ActionResult> EditOnLine(string id)
        {
            var editUrl = await GraphHelper.GetFileEditUrl(id);
            if (editUrl != null)
            {
                return Redirect(editUrl);
            }
            else
            {
                return RedirectToAction("Index");
            }
        }

        [Authorize]
        public async Task<ActionResult> NewDoc(string name)
        {
            var editUrl = await GraphHelper.CreateDocument(name);
            if (editUrl != null)
            {
                return Redirect(editUrl);
            }
            else
            {
                return RedirectToAction("Index");
            }
        }


        [Authorize]
        public async Task<ActionResult> NewSS(string name)
        {
            var editUrl = await GraphHelper.CreateSpreadsheet(name);
            if (editUrl != null)
            {
                return Redirect(editUrl);
            }
            else
            {
                return RedirectToAction("Index");
            }
        }
    }
}